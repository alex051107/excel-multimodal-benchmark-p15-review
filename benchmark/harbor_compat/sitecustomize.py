"""Narrow Harbor 0.22.0 compatibility shim for preinstalled Qwen Code.

Harbor 0.22.0 already skips Codex and Claude Code installation when the exact
requested version exists in the task image. Its Qwen adapter does not. This
shim adds the same exact-version check without modifying Harbor's installed
files. It is loaded only by the public P15 runner through ``PYTHONPATH``.
"""

from __future__ import annotations


def _patch_qwen_install() -> None:
    try:
        from harbor.agents.installed.qwen_code import QwenCode
    except ImportError:
        return

    if getattr(QwenCode, "_p15_preinstalled_version_check", False):
        return

    original_install = QwenCode.install

    async def install(self, environment):  # type: ignore[no-untyped-def]
        result = await environment.exec(
            command="if [ -s ~/.nvm/nvm.sh ]; then . ~/.nvm/nvm.sh; fi; qwen --version"
        )
        lines = (result.stdout or "").strip().splitlines()
        installed = lines[-1].strip() if lines else ""
        requested = getattr(self, "_version", None)
        if result.return_code == 0 and (requested is None or installed == requested):
            self.logger.debug(
                "Qwen Code is already available at the requested version"
            )
            return
        await original_install(self, environment)

    QwenCode.install = install
    QwenCode._p15_preinstalled_version_check = True


_patch_qwen_install()
