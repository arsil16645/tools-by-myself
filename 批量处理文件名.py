import os

# 指定要处理的目录路径（例如：'C:/Users/YourName/Documents/Subs'）
directory = r'H:\86 -不存在的战区[UHA-WINGS&VCB-Studio] EIGHTY SIX [Ma10p_1080p]'  # 替换为你的实际路径

# 获取该目录下所有文件
files = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]

for filename in files:
    if 'sc.ass.txt' in filename:
        old_path = os.path.join(directory, filename)
        new_name = filename.replace('sc.ass.txt', 'sc.ass')
        new_path = os.path.join(directory, new_name)

        os.rename(old_path, new_path)
        print(f'Renamed: {filename} -> {new_name}')

print("\n批量重命名完成！")