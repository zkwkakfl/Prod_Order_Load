# -*- coding: utf-8 -*-
"""
프로젝트 공용 릴리즈(커밋/푸시/태그) 도구.

다른 프로젝트에도 복사해서 그대로 쓸 수 있도록, 외부 라이브러리 없이 표준 라이브러리만 사용한다.
"""

from .core import ReleaseManager, ReleaseOptions

__all__ = ["ReleaseManager", "ReleaseOptions"]

