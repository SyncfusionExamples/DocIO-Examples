import zipfile
import os
import shutil

base_docx = '/workspace/docio-examples/Bookmarks/Add-bookmark-in-Word-document/.NET/Add-bookmark-in-Word-document/Output/Result.docx'
target_dir = '/workspace/docio-examples'
target_file = os.path.join(target_dir, 'document_150mb.docx')
os.makedirs(target_dir, exist_ok=True)

with zipfile.ZipFile(base_docx, 'r') as zin:
    entries = {item.filename: zin.read(item.filename) for item in zin.infolist()}

# Let's update [Content_Types].xml to register bin or dat or png extension
content_types = entries['[Content_Types].xml'].decode('utf-8')
if 'Extension="bin"' not in content_types:
    content_types = content_types.replace('</Types>', '<Default Extension="bin" ContentType="application/octet-stream" /></Types>')
entries['[Content_Types].xml'] = content_types.encode('utf-8')

# Target exactly 150 MB = 150 * 1024 * 1024 = 157,286,400 bytes
TARGET_TOTAL_SIZE = 150 * 1024 * 1024

# First write with a dummy 0-byte payload to compute overhead
with zipfile.ZipFile('/tmp/test_size.zip', 'w') as ztest:
    for name, data in entries.items():
        ztest.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
    ztest.writestr('word/media/embedded_data.bin', b'', compress_type=zipfile.ZIP_STORED)

overhead = os.path.getsize('/tmp/test_size.zip')
payload_size = TARGET_TOTAL_SIZE - overhead

# Now build the exact file
with zipfile.ZipFile(target_file, 'w') as zout:
    for name, data in entries.items():
        zout.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
    
    # Write payload_size bytes
    # Write in chunks using urandom or pseudorandom
    chunk_size = 10 * 1024 * 1024
    remaining = payload_size
    
    zinfo = zipfile.ZipInfo('word/media/embedded_data.bin')
    zinfo.compress_type = zipfile.ZIP_STORED
    
    with zout.open(zinfo, 'w') as f_entry:
        while remaining > 0:
            to_write = min(remaining, chunk_size)
            f_entry.write(os.urandom(to_write))
            remaining -= to_write

actual_size = os.path.getsize(target_file)
print(f"Created {target_file}")
print(f"Target size: {TARGET_TOTAL_SIZE} bytes")
print(f"Actual size: {actual_size} bytes ({actual_size / (1024*1024):.2f} MB)")
