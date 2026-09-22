import os

print("os.path.exists('G:'):", os.path.exists('G:'))
print("os.path.exists('G:\\\\'):", os.path.exists('G:\\\\'))

if os.path.exists('G:\\\\'):
    try:
        items = os.listdir('G:\\\\')
        print("Items in G:\\:", items)
    except Exception as e:
        print("Error listing G:\\:", e)
