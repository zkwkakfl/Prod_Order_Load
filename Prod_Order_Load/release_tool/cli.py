# -*- coding: utf-8 -*-
"""
CLI 엔트리.

사용:
  python -m release_tool.cli --remote auto
  python -m release_tool.cli --remote prod-order
"""

from __future__ import annotations

import argparse
import sys
import subprocess
from pathlib import Path

from .core import ReleaseManager, ReleaseOptions, ReleaseError


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="GitHub 저장 + VERSION 기반 태깅 자동화 (재사용 가능한 모듈)")
    p.add_argument("--project-dir", default="", help="프로젝트 폴더 경로. 비우면 이 CLI 파일의 상위 폴더 기준")
    p.add_argument("--remote", default="auto", help="원격 이름. 기본값: auto(프로젝트 폴더명으로 추정)")
    p.add_argument("--commit-message", default="", help="커밋 메시지. 비우면 기본 메시지 사용")
    p.add_argument("--no-tag", action="store_true", help="태그 생성/푸시를 하지 않음")
    p.add_argument("--allow-empty", action="store_true", help="변경사항이 없어도 진행(기본은 변경 없으면 중단)")
    p.add_argument("--version-file", default="VERSION", help="버전 파일명(기본: VERSION)")
    return p


def main(argv: list[str]) -> int:
    args = build_parser().parse_args(argv)

    if args.project_dir:
        project_dir = Path(args.project_dir).resolve()
    else:
        # 기본: 이 모듈이 들어있는 프로젝트 폴더의 한 단계 위를 프로젝트 폴더로 가정하지 않고,
        # 사용자가 보통 "프로젝트 폴더에 복사해 실행"한다는 전제에서, 현재 작업 디렉터리 기준이 더 안전하다.
        project_dir = Path.cwd().resolve()

    opt = ReleaseOptions(
        remote=args.remote,
        commit_message=args.commit_message,
        no_tag=bool(args.no_tag),
        allow_empty=bool(args.allow_empty),
        version_file=args.version_file,
    )
    mgr = ReleaseManager(project_dir=project_dir, options=opt)

    try:
        res = mgr.run_release()
        # stdout 메시지는 스크립트/CI에서도 파싱하기 쉽도록 단순 텍스트로 유지
        print(res.get("message", ""))
        print(f"remote={res.get('remote','')}")
        if res.get("remote_url"):
            print(f"remote_url={res.get('remote_url','')}")
        print(f"branch={res.get('branch','')}")
        if res.get("tag"):
            print(f"tag={res.get('tag')}")
        return 0 if res.get("status") in ("ok", "noop") else 1
    except ReleaseError as e:
        print(f"오류: {e}")
        return 2
    except subprocess.CalledProcessError as e:
        # core에서 subprocess를 직접 다루지만, 혹시 모를 예외를 대비
        print("명령 실행 실패")
        print(str(e))
        return 1
    except Exception as e:
        print(f"오류: {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

