import os
import psse35
import dyntools
import pandas as pd

def checker(file, name_file, result,bus):
    chnf = dyntools.CHNF(file)
    short_tiltle, channel_ids, channel_data = chnf.get_data()

    bus = str(bus)
    k1 = 0
    for key, value in channel_ids.items():
        if "POWR" in value and bus in value:
            k1 = key
            break
    a = channel_data[k1][0]
    b = channel_data[k1][-1]
    if abs((1 - b / a) * 100) >= 4:
        result[name_file] = abs((1 - b / a) * 100)

def process_files(folder, except_folders,bus):
    result = {}
    path = os.path.join(folder, "RESULTs")
    for root, dirs, files in os.walk(path):
        dirs[:] = [d for d in dirs if d not in except_folders]
        for file in files:
            if file.endswith('.out'):
                try:
                    checker(os.path.join(root, file), file, result,bus)
                except Exception as e:
                    print(f"Error processing file {file}: {e}")
    return result

def save_to_excel(result, folder):
    parent_folder = os.path.dirname(folder)  # Lấy thư mục cha

    filename = os.path.join(parent_folder,"P_drop.xlsx")
    df = pd.DataFrame(list(result.items()), columns=['File Name', 'P drop'])
    print(filename)
    df.to_excel(filename, index=False)

if __name__ == "__main__":
    folder = r"D:\WORK\Bandana Storage\MQT_autotap\5PV_BESSD"
    bus =100003
    except_folders = ['2PROJECT_LAG_HV_PREFER_PVBESS2', '2PROJECT_LEAD_HV_PREFER_PVBESS2']
    result = process_files(folder, except_folders,bus)
    print(result)
    save_to_excel(result,folder)