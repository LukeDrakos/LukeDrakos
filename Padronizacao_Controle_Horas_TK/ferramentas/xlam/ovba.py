import struct
def _bitcount(pos):
    # pos = bytes already decompressed in this chunk (>0)
    b = 4
    while (1 << b) < pos: b += 1
    return b

def compress_chunk(data):
    out = bytearray(); i = 0; n = len(data)
    while i < n:
        flag_pos = len(out); out.append(0); flag = 0
        for bit in range(8):
            if i >= n: break
            # procura a maior repeticao dentro do chunk
            best_len = 0; best_off = 0
            if i > 0:
                bc = _bitcount(i)
                max_len = (0xFFFF >> bc) + 3
                lim = min(max_len, n - i)
                start = max(0, i - (1 << bc))
                for j in range(i - 1, start - 1, -1):
                    L = 0
                    while L < lim and data[j + L] == data[i + L]:
                        L += 1
                    if L > best_len:
                        best_len, best_off = L, i - j
                        if L == lim: break
            if best_len >= 3:
                bc = _bitcount(i)
                tok = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                out += struct.pack('<H', tok); flag |= (1 << bit); i += best_len
            else:
                out.append(data[i]); i += 1
        out[flag_pos] = flag
    return bytes(out)

def compress(data):
    out = bytearray(b'\x01')
    for k in range(0, len(data), 4096):
        chunk = data[k:k + 4096]
        comp = compress_chunk(chunk)
        if len(comp) + 2 > 4098 and len(chunk) == 4096:
            out += struct.pack('<H', (4096 + 2 - 3) | 0x3000)  # chunk sem compressao
            out += chunk
        else:
            out += struct.pack('<H', (len(comp) + 2 - 3) | 0x3000 | 0x8000)
            out += comp
    return bytes(out)
