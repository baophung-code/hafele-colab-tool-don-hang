# -*- coding: utf-8 -*-
"""So khớp mờ địa chỉ — copy nguyên logic đã verify từ scripts/hafele_lib.py (qua 7 lần sửa
bug trên data thật, xem docs/QUYLUAT_GOM_BBGH.md §6). KHÔNG sửa logic ở đây.

Tách riêng module này vì hafele_lib.py có import thư viện pandas ở đầu file — import cả module
đó vào scripts/tool/ sẽ kéo pandas vào exe (exe chỉ dùng openpyxl). Các hàm dưới đây không dùng
pandas.
"""
import re
import unicodedata


def strip_accent(s):
    s = unicodedata.normalize('NFD', s)
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn').replace('đ', 'd').replace('Đ', 'D')


_RUN = re.compile(r'\d[\d\s\.\-\(\)]*\d|\d')
_NONWORD = re.compile(r'[^0-9a-z]+')

STOP = set("""phuong p quan q huyen h xa x thi tran tt tp thanh pho tinh tx duong d so no
khu cn kcn lo lot to ap khom block toa nha building floor lau tang phong room
viet nam vietnam vn cong ty tnhh cp co phan chi nhanh mtv sdt dt lh lien he
giao hang nhan nguoi kho tel mobile phone anh chi ong ba mr ms mrs a c ng""".split())


def parse_address(raw):
    """-> (list SĐT, set token địa chỉ)."""
    if not isinstance(raw, str):
        return [], set()
    phones = []

    def take(m):
        run = m.group(0)
        s = re.sub(r'\D', '', run)
        found, nums, i, last = [], [], 0, 0
        while i <= len(s) - 9:
            if s[i] == '0' and s[i + 1] in '2356789':
                rest = len(s) - i
                L = 11 if (s[i + 1] == '2' and rest == 11) else (10 if rest >= 10 else rest)
                if last < i:
                    nums.append(s[last:i])
                found.append(s[i:i + L])
                i = last = i + L
                continue
            i += 1
        if last < len(s):
            nums.append(s[last:])
        if not found:
            return run
        phones.extend(found)
        return ' ' + ' '.join(nums) + ' '

    t = _RUN.sub(take, raw)
    toks = set()
    for x in _NONWORD.split(strip_accent(t).lower()):
        for y in _split_glued(x):
            if not y or y in STOP:
                continue
            if len(y) == 1 and not y.isdigit():
                continue
            if y.isdigit():
                y = y.lstrip('0') or '0'
            toks.add(y)
    return list(dict.fromkeys(phones)), toks


_GLUED = re.compile(r'^(\d+)([a-z]{2,})$|^([a-z]{2,})(\d+)$')


def _split_glued(x):
    m = _GLUED.match(x)
    if not m:
        return (x,)
    g = [p for p in m.groups() if p]
    return tuple(g)


def jaccard(a, b):
    return len(a & b) / len(a | b) if a and b else 0.0


def contain(a, b):
    return len(a & b) / min(len(a), len(b)) if a and b else 0.0


def num_tokens(toks):
    return {t for t in toks if any(c.isdigit() for c in t)}


def _num_compat(x, y):
    """'15' ~ '15B' (gõ thiếu hậu tố chữ) nhưng '1' ≠ '15' và 'G1' ≠ 'B1'."""
    if x == y:
        return True
    lo, hi = (x, y) if len(x) < len(y) else (y, x)
    return hi.startswith(lo) and hi[len(lo):].isalpha()


def num_conflict(a, b):
    na, nb = num_tokens(a), num_tokens(b)
    if not na or not nb:
        return False
    la = any(not any(_num_compat(x, y) for y in nb) for x in na)
    lb = any(not any(_num_compat(x, y) for y in na) for x in nb)
    return la and lb


def same_place(a, b, jac_min=0.6, cov_min=0.80):
    if num_conflict(a, b):
        return False
    return jaccard(a, b) >= jac_min or contain(a, b) >= cov_min
