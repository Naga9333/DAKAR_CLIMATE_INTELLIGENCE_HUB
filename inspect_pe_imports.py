import pathlib
import struct

path = pathlib.Path(r"C:\Users\pc\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules\@oai\artifact-tool\node_modules\skia-canvas\lib\skia.node")
data = path.read_bytes()
pe = struct.unpack_from("<I", data, 0x3C)[0]
num_sections = struct.unpack_from("<H", data, pe + 6)[0]
optional_size = struct.unpack_from("<H", data, pe + 20)[0]
optional = pe + 24
magic = struct.unpack_from("<H", data, optional)[0]
data_dir = optional + (112 if magic == 0x20B else 96)
import_rva, _ = struct.unpack_from("<II", data, data_dir + 8)
section_start = optional + optional_size
sections = []
for i in range(num_sections):
    offset = section_start + (40 * i)
    virtual_size, virtual_address, raw_size, raw_pointer = struct.unpack_from("<IIII", data, offset + 8)
    sections.append((virtual_address, max(virtual_size, raw_size), raw_pointer))

def to_offset(rva):
    for virtual_address, size, raw_pointer in sections:
        if virtual_address <= rva < virtual_address + size:
            return raw_pointer + (rva - virtual_address)
    return None

import_offset = to_offset(import_rva)
names = []
i = 0
while import_offset is not None:
    descriptor = data[import_offset + (20 * i):import_offset + (20 * (i + 1))]
    if len(descriptor) < 20 or descriptor == b"\0" * 20:
        break
    name_rva = struct.unpack_from("<I", descriptor, 12)[0]
    name_offset = to_offset(name_rva)
    if name_offset:
        end = data.index(b"\0", name_offset)
        names.append(data[name_offset:end].decode("ascii", "ignore"))
    i += 1

print("\n".join(names))
