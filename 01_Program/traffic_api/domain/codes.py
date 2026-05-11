from __future__ import annotations


def normalize_drct_cd(drct_cd) -> str:
    if drct_cd is None:
        return ""
    return str(drct_cd).strip().zfill(2)


def normalize_vknd_cd(vknd_cd) -> str:
    if vknd_cd is None:
        return ""
    return str(vknd_cd).strip()

