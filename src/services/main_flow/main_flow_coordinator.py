from __future__ import annotations

from pathlib import Path
import time
from typing import Any, Callable

from .main_flow_models import (
    MainFlowCommand,
    MainFlowContext,
    SingleArticleOptions,
    SingleArticleReceipt,
)


class MainFlowCoordinator:
    """把主页扫描和单篇任务按安全顺序串联起来。"""

    def __init__(
        self,
        *,
        scanner: Any,
        dispatcher: Any,
        sleep: Callable[[float], None] = time.sleep,
        foreground_timeout_seconds: float = 120.0,
        post_task_timeout_seconds: float = 300.0,
        trace: Callable[[dict[str, Any]], None] | None = None,
    ) -> None:
        self._scanner = scanner
        self._dispatcher = dispatcher
        self._sleep = sleep
        self._foreground_timeout_seconds = max(0.1, float(foreground_timeout_seconds))
        self._post_task_timeout_seconds = max(0.1, float(post_task_timeout_seconds))
        self._trace = trace

    def run(self, context: MainFlowContext | Any, command: MainFlowCommand) -> None:
        state = context.state
        handles: list[Any] = []
        handled_fingerprints: set[str] = set()
        session: Any | None = None
        flow_error: Exception | None = None
        article_error_reported = False
        try:
            if str(state.snapshot().get("status")) in {"idle", "starting"}:
                state.start()
            self._set_action(state, "正在定位公众号主页")
            session = self._scanner.initialize(command)
            account_name = str(getattr(session, "account_name", "") or "").strip()
            if account_name:
                state.set_account_name(account_name)

            while not context.cancel_token.is_set():
                if _target_limit_reached(state, command.target_count):
                    break
                self._set_action(state, "正在识别文章卡片")
                batch = self._scanner.scan_visible_targets(session)
                dispatched = False
                for target in tuple(getattr(batch, "targets", ()) or ()):
                    if context.cancel_token.is_set():
                        break
                    if target.fingerprint in handled_fingerprints:
                        continue
                    if _target_limit_reached(state, command.target_count):
                        break
                    handled_fingerprints.add(target.fingerprint)
                    state.set_task_info(target.title_raw or target.title_display)
                    options = SingleArticleOptions(
                        skip_collected_records=command.skip_collected_records,
                        collect_article_detail=True,
                        collect_comments=command.collect_comments,
                        archive_offline=command.archive_offline,
                        offline_archive_mode=command.offline_archive_mode,
                    )
                    self._set_action(state, "正在分发单篇任务")
                    handle = self._dispatcher.submit(target, options, context)
                    handles.append(handle)
                    dispatched = True
                    self._set_action(state, "正在等待文章标签关闭")
                    try:
                        foreground = self._wait_foreground_done(
                            handle,
                            context=context,
                        )
                    except Exception as exc:
                        state.handle_receipt(
                            SingleArticleReceipt(
                                task_id=str(getattr(handle, "handle_id", "")),
                                target_fingerprint=target.fingerprint,
                                status="failed",
                                phase="foreground",
                                error_stage="foreground",
                                error_detail=str(exc),
                                message="前台阶段未返回安全回执",
                            )
                        )
                        article_error_reported = True
                        self._trace_event("foreground-timeout", target=target, error=str(exc))
                        raise

                    state.handle_receipt(foreground)
                    self._drain_updates(handle, state)
                    self._scanner.mark_processed(session, target)
                    if not foreground.foreground_done:
                        raise RuntimeError("单篇任务未确认主页可以继续")
                    if command.single_task_interval_seconds > 0:
                        self._sleep(command.single_task_interval_seconds)

                if context.cancel_token.is_set() or _target_limit_reached(
                    state, command.target_count
                ):
                    break
                if bool(getattr(batch, "boundary_reached", False)):
                    break
                self._set_action(state, "正在滚动主页")
                if not self._scanner.scroll_next(session):
                    break
                if not dispatched:
                    # 空屏仍然允许扫描器执行下一次滚动；不在协调器中读取 UIA。
                    continue

        except Exception as exc:
            flow_error = exc
            self._trace_event("main-flow-failed", error=str(exc))
        finally:
            self._set_action(state, "正在清理运行资源")
            should_cancel = flow_error is not None or context.cancel_token.is_set()
            if should_cancel:
                self._cancel_handles(handles)
            error_count_before = int(state.snapshot().get("errorCount") or 0)
            self._wait_post_tasks(handles, state, context)
            if context.cancel_token.is_set():
                state.cancel()
            elif flow_error is not None:
                error_count_after = int(state.snapshot().get("errorCount") or 0)
                state.fail(
                    f"主流程异常：{type(flow_error).__name__}: {flow_error}",
                    count_error=(
                        not article_error_reported
                        and error_count_after == error_count_before
                    ),
                )
            elif state.snapshot().get("status") not in {"failed", "cancelled"}:
                state.complete()

    def _wait_post_tasks(self, handles: list[Any], state: Any, context: Any) -> None:
        for handle in handles:
            try:
                final = self._dispatcher.wait_finished(
                    handle,
                    timeout_seconds=self._post_task_timeout_seconds,
                )
                state.handle_receipt(final)
                self._drain_updates(handle, state)
            except Exception as exc:
                target = getattr(handle, "target", None)
                fingerprint = str(getattr(target, "fingerprint", "") or "")
                if fingerprint:
                    state.handle_receipt(
                        SingleArticleReceipt(
                            task_id=str(getattr(handle, "handle_id", "")),
                            target_fingerprint=fingerprint,
                            status="failed",
                            phase="finished",
                            error_stage="post_task",
                            error_detail=str(exc),
                            message="后置任务等待失败",
                        )
                    )
                if context.cancel_token.is_set():
                    self._dispatcher.cancel(handle)

    def _cancel_handles(self, handles: list[Any]) -> None:
        cancel = getattr(self._dispatcher, "cancel", None)
        if not callable(cancel):
            return
        for handle in handles:
            try:
                cancel(handle)
            except Exception:
                continue

    def _wait_foreground_done(
        self,
        handle: Any,
        *,
        context: MainFlowContext | Any,
    ) -> Any:
        """以短轮询等待前台回执，使用户停止可以及时打断等待。"""

        deadline = time.monotonic() + self._foreground_timeout_seconds
        while True:
            if context.cancel_token.is_set():
                self._cancel_handles([handle])
                raise RuntimeError("主流程已请求停止")
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise TimeoutError("单篇任务未在规定时间返回前台回执")
            try:
                return self._dispatcher.wait_foreground_done(
                    handle,
                    timeout_seconds=min(0.2, remaining),
                )
            except TimeoutError:
                if context.cancel_token.is_set():
                    self._cancel_handles([handle])
                    raise RuntimeError("主流程已请求停止")
                if time.monotonic() >= deadline:
                    raise

    def _drain_updates(self, handle: Any, state: Any) -> None:
        poll = getattr(self._dispatcher, "poll_updates", None)
        updates = poll(handle) if callable(poll) else list(getattr(handle, "updates", ()) or ())
        for update in updates:
            receipt = getattr(update, "to_receipt", None)
            if callable(receipt):
                state.handle_receipt(receipt())

    def _set_action(self, state: Any, action: str) -> None:
        state.set_action(action)
        self._trace_event("action", action=action)

    def _trace_event(self, event: str, **details: Any) -> None:
        if self._trace is None:
            return
        try:
            self._trace({"event": event, "details": details})
        except Exception:
            return


def _target_limit_reached(state: Any, target_count: int) -> bool:
    if int(target_count) <= 0:
        return False
    snapshot = state.snapshot()
    return int(snapshot.get("progressDone") or 0) >= int(target_count)


__all__ = ["MainFlowCoordinator"]
