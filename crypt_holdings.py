# -*- coding: utf-8 -*-
"""
持仓文件加解密工具（stdlib only，供本地加密与 Actions 解密共用）。
用法：
  python crypt_holdings.py encrypt <输入文件> <输出文件> <hexkey>
  python crypt_holdings.py decrypt <输入文件> <输出文件> <hexkey>
格式：MAGIC(8) + [HMAC-SHA256 keystream XOR]
"""
import hashlib
import hmac
import sys

MAGIC = b"FOFENC01"


def keystream(key_bytes, length):
    out = bytearray()
    counter = 0
    while len(out) < length:
        out.extend(hmac.new(key_bytes, b"blk%d" % counter, hashlib.sha256).digest())
        counter += 1
    return bytes(out[:length])


def xor_bytes(data, key_bytes):
    ks = keystream(key_bytes, len(data))
    return bytes(a ^ b for a, b in zip(data, ks))


def main():
    mode, src, dst, key = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
    key_bytes = bytes.fromhex(key)
    with open(src, "rb") as f:
        data = f.read()
    if mode == "encrypt":
        raw = data
        if raw.startswith(MAGIC):
            sys.exit("already encrypted")
        out = MAGIC + xor_bytes(raw, key_bytes)
    else:
        if not data.startswith(MAGIC):
            sys.exit("not an encrypted file")
        out = xor_bytes(data[len(MAGIC):], key_bytes)
    with open(dst, "wb") as f:
        f.write(out)
    print("%s OK: %s -> %s (%d bytes)" % (mode, src, dst, len(out)))


if __name__ == "__main__":
    main()
