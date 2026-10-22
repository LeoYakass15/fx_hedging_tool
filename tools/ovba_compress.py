"""MS-OVBA 2.4.1 compression (LZ77 variant used for VBA streams), written from the spec.

Replaces the compressor of the ms_ovba_compression package, which corrupts inputs
longer than ~3 KB. Round-trip tested against the oletools decompressor.
"""
from __future__ import annotations

CHUNK = 4096


def _copy_token_help(difference: int):
    bit_count = max((difference - 1).bit_length(), 4)
    length_mask = 0xFFFF >> bit_count
    return bit_count, length_mask, length_mask + 3   # max match length


def _compress_chunk(chunk: bytes) -> bytes:
    out = bytearray()
    pos, n = 0, len(chunk)
    index: dict[bytes, list[int]] = {}
    while pos < n:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            best_len, best_off = 0, 0
            if pos >= 1 and pos + 3 <= n:
                bit_count, length_mask, max_len = _copy_token_help(pos)
                for cand in reversed(index.get(chunk[pos:pos + 3], [])):
                    ln = 0
                    limit = min(max_len, n - pos)
                    while ln < limit and chunk[cand + ln] == chunk[pos + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, pos - cand
                        if ln == limit:
                            break
            if best_len >= 3:
                bit_count, length_mask, _ = _copy_token_help(pos)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += token.to_bytes(2, "little")
                flags |= 1 << bit
                step = best_len
            else:
                out.append(chunk[pos])
                step = 1
            for k in range(pos, pos + step):
                if k + 3 <= n:
                    index.setdefault(chunk[k:k + 3], []).append(k)
            pos += step
        out[flag_pos] = flags
    return bytes(out)


def compress(data: bytes) -> bytes:
    container = bytearray([0x01])
    for start in range(0, len(data), CHUNK):
        chunk = data[start:start + CHUNK]
        body = _compress_chunk(chunk)
        if len(body) + 2 > 4098:
            if len(chunk) != CHUNK:
                raise ValueError("incompressible final chunk")
            container += (0x3000 | 4095).to_bytes(2, "little") + chunk
        else:
            header = ((len(body) + 2 - 3) & 0x0FFF) | 0xB000   # signature 011, flag 1
            container += header.to_bytes(2, "little") + body
    return bytes(container)


def patch_ms_ovba() -> None:
    from ms_ovba_compression.ms_ovba import MsOvba
    MsOvba.compress = lambda self, data: compress(bytes(data))
