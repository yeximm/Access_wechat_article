from __future__ import annotations

import dev_server


def main() -> None:
    """兼容旧启动命令；真正的 API 后端入口统一放在 dev_server.py。"""
    dev_server.main()


if __name__ == "__main__":
    main()
