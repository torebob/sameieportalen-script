partition_table = """\
# Navn,     Type, SubType, Offset,  Size, Flags
nvs,        data, nvs,     0x9000,  0x5000,
otadata,    data, ota,     0xe000,  0x2000,
app0,       app,  ota_0,   0x10000, 0x140000,
app1,       app,  ota_1,   0x150000,0x140000,
spiffs,     data, spiffs,  0x290000,0x170000,
"""

filename = "partitions.csv"

with open(filename, "w") as f:
    f.write(partition_table)

print(f"{filename} er opprettet!")
