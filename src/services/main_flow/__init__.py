"""主服务任务编排模块。

该包只负责主流程的命令、状态和任务生命周期；窗口、代理、解析和存储能力
由后续接入的独立模块提供。
"""

from .main_flow_models import (
    HomeArticleTarget,
    MainFlowCommand,
    MainFlowContext,
    MainFlowSnapshot,
    SingleArticleOptions,
    SingleArticleReceipt,
)
from .main_flow_service import MainFlowConflictError, MainFlowService
from .main_flow_state import MainFlowState
from .main_flow_factory import build_main_flow_service

__all__ = [
    "HomeArticleTarget",
    "MainFlowCommand",
    "MainFlowConflictError",
    "MainFlowContext",
    "MainFlowService",
    "MainFlowSnapshot",
    "MainFlowState",
    "SingleArticleOptions",
    "SingleArticleReceipt",
    "build_main_flow_service",
]
