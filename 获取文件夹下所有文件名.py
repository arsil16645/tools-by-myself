import os
def first_dir():
    #F:\\二区\\game
    #F:\\二区\\video

    # 指定目录路径
    directory_path = 'F:\\二区\\game'
    txt_file = f'{directory_path}\\name.txt'
    # 获取指定目录下所有文件和文件夹的名称
    file_names = [f for f in os.listdir(directory_path) if
                  os.path.isfile(os.path.join(directory_path, f)) or os.path.isdir(os.path.join(directory_path, f))]
    print(file_names)
    for i in file_names:
        with open(txt_file, 'a', encoding='utf-8') as f:
            f.write(f"{i}\n")
    print("ok")

def sec_dir():
    #F:\\一区\\video
    #F:\\一区\\game
    # F:\\二区\\video\\B站

    # 指定目录路径
    directory_path = 'F:\\一区\\video'
    txt_file = f'{directory_path}\\name.txt'
    # 获取指定目录下所有文件和文件夹的名称，包括文件夹里面的文件名
    file_names = []
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            file_names.append(file)
    print(file_names)
    for i in file_names:
        with open(txt_file, 'a', encoding='utf-8') as f:
            f.write(f"{i}\n")
    print("ok")

if __name__=="__main__":
    #first_dir()

    sec_dir()