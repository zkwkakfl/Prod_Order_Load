# -*- coding: utf-8 -*-
"""
ReleaseManager: 커밋/푸시/태깅 코어 로직.

설계 포인트
- "프로젝트 폴더" 단위로만 스테이징하는 기본값으로 안전하게 동작
- 원격 선택은 기본(auto)로 "프로젝트 폴더명 == 원격 repo명" 매칭
- VERSION 파일(SemVer)로 태그(vX.Y.Z) 생성

빠른 사용법

1) 코드에서 직접 사용(재사용 목적)

    from pathlib import Path
    from release_tool.core import ReleaseManager, ReleaseOptions

    opt = ReleaseOptions(remote="auto")  # 또는 remote="prod-order"
    mgr = ReleaseManager(project_dir=Path.cwd(), options=opt)
    result = mgr.run_release()
    print(result)

2) 기본 동작 요약
- 스테이징 범위: project_dir 폴더 내부만 (상위 폴더의 다른 프로젝트가 섞이는 것을 방지)
- 커밋 메시지: options.commit_message가 비어 있으면 "chore(release): prepare vX.Y.Z"
- 푸시 브랜치: 현재 브랜치가 master/main이면 원격 main으로 푸시(HEAD:main)
- 태그: VERSION이 1.4.0 이면 v1.4.0 태그(annotated) 생성 후 원격에 푸시

3) 예외/주의
- 이미 스테이징된 변경사항이 있으면 중단(ReleaseError)
- 태그가 로컬/원격에 이미 존재하면 중단(ReleaseError)
"""

from __future__ import annotations

import os
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path


SEMVER_RE = re.compile(r"^\d+\.\d+\.\d+$")


class ReleaseError(RuntimeError):
    pass


@dataclass(frozen=True)
class ReleaseOptions:
    remote: str = "auto"
    commit_message: str = ""
    no_tag: bool = False
    allow_empty: bool = False
    version_file: str = "VERSION"


class ReleaseManager:
    def __init__(self, *, project_dir: Path, options: ReleaseOptions | None = None):
        self.project_dir = project_dir.resolve()
        self.options = options or ReleaseOptions()
        self.repo_root = self._git_root(self.project_dir)
        self.project_rel = os.path.relpath(str(self.project_dir), str(self.repo_root))
        self.project_name = self.project_dir.name

    # -------------------------
    # Git helpers
    # -------------------------
    def _run(self, cmd: list[str], *, cwd: Path | None = None, check: bool = True) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            cmd,
            cwd=str((cwd or self.repo_root)),
            text=True,
            capture_output=True,
            check=check,
            encoding="utf-8",
            errors="replace",
        )

    @staticmethod
    def _git_root(start: Path) -> Path:
        cp = subprocess.run(
            ["git", "rev-parse", "--show-toplevel"],
            cwd=str(start),
            text=True,
            capture_output=True,
            check=True,
            encoding="utf-8",
            errors="replace",
        )
        return Path(cp.stdout.strip()).resolve()

    def current_branch(self) -> str:
        return self._run(["git", "rev-parse", "--abbrev-ref", "HEAD"]).stdout.strip()

    def list_remotes(self) -> dict[str, str]:
        """
        remote_name -> push_url
        """
        cp = self._run(["git", "remote", "-v"])
        out: dict[str, str] = {}
        for line in cp.stdout.splitlines():
            parts = line.strip().split()
            if len(parts) != 3:
                continue
            name, url, kind = parts
            if kind != "(push)":
                continue
            out[name] = url
        return out

    @staticmethod
    def guess_remote_by_project_name(remotes: dict[str, str], project_name: str) -> str | None:
        target = project_name.casefold()
        for name, url in remotes.items():
            m = re.search(r"/([^/]+?)(?:\.git)?$", url)
            if not m:
                continue
            repo = m.group(1).casefold()
            if repo == target:
                return name
        return None

    def resolve_remote(self) -> tuple[str, dict[str, str]]:
        remotes = self.list_remotes()
        if not remotes:
            raise ReleaseError("원격(remote)이 설정되어 있지 않습니다. git remote -v 를 확인하세요.")

        remote_opt = (self.options.remote or "").strip()
        if remote_opt == "auto":
            guessed = self.guess_remote_by_project_name(remotes, self.project_name)
            if not guessed:
                raise ReleaseError(
                    "auto 원격 선택 실패\n"
                    f"- 프로젝트 폴더명: {self.project_name}\n"
                    f"- 등록된 원격: {', '.join(remotes.keys())}\n"
                    "해결: --remote prod-order 처럼 원격 이름을 직접 지정하세요."
                )
            return guessed, remotes

        if remote_opt not in remotes:
            raise ReleaseError(f"원격 {remote_opt!r} 이(가) 없습니다. 등록된 원격: {', '.join(remotes.keys())}")
        return remote_opt, remotes

    def read_version(self) -> str:
        p = self.project_dir / self.options.version_file
        if not p.is_file():
            raise ReleaseError(f"VERSION 파일이 없습니다: {p}")
        v = (p.read_text(encoding="utf-8", errors="replace") or "").strip()
        if not SEMVER_RE.fullmatch(v):
            raise ReleaseError(f"VERSION 형식이 SemVer(예: 1.4.0)가 아닙니다: {v!r}")
        return v

    def tag_name(self) -> str:
        return f"v{self.read_version()}"

    def ensure_clean_index(self) -> None:
        cp = self._run(["git", "diff", "--cached", "--name-only"])
        if cp.stdout.strip():
            raise ReleaseError("이미 스테이징된 변경사항이 있습니다. 먼저 정리 후 다시 실행하세요.")

    def project_status_porcelain(self) -> str:
        cp = self._run(["git", "status", "--porcelain=v1", "--", self.project_rel])
        return cp.stdout

    def stage_project_only(self) -> None:
        self._run(["git", "add", "-A", "--", self.project_rel])

    def has_staged_changes(self) -> bool:
        cp = self._run(["git", "diff", "--cached", "--name-only"])
        return bool(cp.stdout.strip())

    def commit(self, message: str) -> None:
        self._run(["git", "commit", "-m", message])

    def push_branch(self, remote: str, branch: str) -> None:
        # 원격 기본 브랜치가 main인 경우가 많아서 master/main은 remote/main으로 통일.
        if branch in ("master", "main"):
            self._run(["git", "push", remote, "HEAD:main"])
            return
        self._run(["git", "push", remote, f"HEAD:{branch}"])

    def tag_exists_local(self, tag: str) -> bool:
        return bool(self._run(["git", "tag", "--list", tag]).stdout.strip())

    def tag_exists_remote(self, remote: str, tag: str) -> bool:
        return bool(self._run(["git", "ls-remote", "--tags", remote, f"refs/tags/{tag}"]).stdout.strip())

    def create_and_push_tag(self, remote: str, tag: str) -> None:
        self._run(["git", "tag", "-a", tag, "-m", f"release: {tag}"])
        self._run(["git", "push", remote, tag])

    # -------------------------
    # High-level orchestration
    # -------------------------
    def run_release(self) -> dict[str, str]:
        """
        변경분 커밋/푸시 후 태깅까지 수행.
        반환: 수행 결과 요약 dict
        """
        remote, remotes = self.resolve_remote()
        tag = self.tag_name()
        branch = self.current_branch()

        self.ensure_clean_index()

        st = self.project_status_porcelain()
        if not st.strip() and not self.options.allow_empty:
            return {
                "status": "noop",
                "remote": remote,
                "remote_url": remotes.get(remote, ""),
                "branch": branch,
                "tag": tag if not self.options.no_tag else "",
                "message": "변경사항이 없습니다(프로젝트 폴더 기준).",
            }

        # 원격 최신 반영
        self._run(["git", "fetch", remote])

        # 커밋
        self.stage_project_only()
        if self.has_staged_changes():
            msg = (self.options.commit_message or "").strip()
            if not msg:
                msg = f"chore(release): prepare {tag}"
            self.commit(msg)
        else:
            if not self.options.allow_empty:
                return {
                    "status": "noop",
                    "remote": remote,
                    "remote_url": remotes.get(remote, ""),
                    "branch": branch,
                    "tag": tag if not self.options.no_tag else "",
                    "message": "스테이징 결과가 비어 있어 커밋을 만들지 않았습니다.",
                }

        # 푸시
        self.push_branch(remote, branch)

        # 태그
        if not self.options.no_tag:
            if self.tag_exists_local(tag) or self.tag_exists_remote(remote, tag):
                raise ReleaseError(f"태그 {tag} 가 이미 존재합니다. 중복 태깅을 중단합니다.")
            self.create_and_push_tag(remote, tag)

        return {
            "status": "ok",
            "remote": remote,
            "remote_url": remotes.get(remote, ""),
            "branch": branch,
            "tag": tag if not self.options.no_tag else "",
            "message": "완료",
        }

