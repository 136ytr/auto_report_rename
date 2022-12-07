import os
import sys
import re
import xlwings as xw
# from colorama import init
# init(autoreset=True)

if os.name == "nt":
    os.system("")

os.system("title 实验报告修改器(Designed By PuMu)")
# path = r"C:\Users\136ytr\OneDrive - 汕头大学\课程\机能学实验\大三上\20221111"
print("使用前请关闭所有Excel窗口，确保主文件命名无误")
path = input("请输入文件夹路径：") or input("不能为空，请重新输入文件夹路径：")

success_num = 0
fail_num = 0
zhuzhang_number = '2020810039'
filelist_data_dict = {}
new_name_list = []
name_list = [   ('FileList', '目录文件', '“%s_机能学实验班级简称_所在实验室名称_实验日期am_学号.xlsx”'), 
                ('主文件', '主文件', '“%s_机能学实验班级简称_所在实验室名称_实验日期am_学号.xlsx”'), 
                ('结果\d+_', '原始实验结果文件', '“%s实验日期am_学号.tme?”'), 
                ('结果\d+Cut', '剪辑后的原始实验结果文件', '“%s_实验日期am_学号.tme?”'),
                ('结果\d+OK', '最终结果文件', '“%s_实验日期am_学号.jpg?”'),
                ('其他\d+_', '其他文件', '“%s实验日期am_学号.jpg?”')
                # ('其他\d{1}_', '其他文件', '“%s实验日期am_学号.jpg?”'),
                # ('其他\d{2}_', '其他文件', '“%s实验日期am_学号.jpg?”')
            ]

app = xw.App(visible=True, add_book=False)

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)
    
def show_all_files(path, all_files = []):
    file_list = os.listdir(path)
    for file in file_list:
        cur_path = os.path.join(path, file)
        if os.path.isdir(cur_path):
            show_all_files(cur_path, all_files)
        else:
            all_files.append(os.path.join(path, file))

    return all_files

all_files_list = show_all_files(path)

print("\n读取主文件中...")
try:
    for file in all_files_list:
        if '主文件' in file:
            file = os.path.basename(file)
            new_name_list.append(file)
            file = os.path.splitext(file)[0]
            _, class_name, classroom, date, number = file.split('_')
            break
    print("班级:", class_name, "\t教室:", classroom, "\t时间:", date, "\t学号:", number)
    print("读取主文件成功")
except Exception as e:
    print("\033[1;31m读取主文件失败\033[0m")
    print(e)
    input('按任意键退出...')
    exit()


print("\n读取并删除FileList文件中...")
try:
    new_name_list.append('FileList_' + class_name + '_' + classroom + '_' + date + '_' + number + '.xlsx')
    for file in all_files_list:
        if 'FileList' in file:
            print("读取FileList文件中...")
            wb = app.books.open(file)
            sht = wb.sheets["sheet1"]
            filelist_data_list = sht.range('D:F').value
            filelist_data_dict = {str(i[0]).split('_')[0]:i[2] for i in filelist_data_list}
            wb.close()
            
            print("删除" + file)
            os.remove(file)
            print("删除FileList文件成功")
            break
    else:
        print("未发现FileList文件")
except Exception as e:
    print("\033[1;31m读取并删除FileList文件失败\033[0m")
    print(e)


print("\n重命名文件中...")
for original_path_name in all_files_list:
    try:
        if '主文件' in original_path_name or 'FileList' in original_path_name:
            continue
        file_path, original_name = os.path.split(original_path_name)
        file_name = original_name.split('_')[0]
        _, file_type = os.path.splitext(original_name)
        new_name = file_name + '_' + date + '_' + number + file_type
        new_name_list.append(new_name)
        new_path_name = os.path.join(file_path, new_name)
        os.rename(original_path_name, new_path_name)
    except:
        print("\033[1;31m处理 " + original_path_name + " 失败\033[0m")
        fail_num += 1
    else:
        print("将 " + original_path_name + " 重命名为 " + new_name)
        success_num += 1
print("重命名文件完成")
print("共处理 " + str(success_num+fail_num) + " 个文件，成功 " + str(success_num) + " 个，\033[1;31m失败 " + str(fail_num) + " 个\033[0m")


print("\n生成FileList文件中...")
try:
    new_name_list_sorted = sorted(new_name_list, key=lambda i: len(i))
    wb = app.books.open(resource_path('FileList_Demo.xlsx'))
    sht = wb.sheets["sheet1"]
    rows = 9
    for match_name, write_name, rule_name in name_list:
        i = 0
        for file in new_name_list_sorted:
            try:
                if re.match(match_name, file):
                    file_name, file_type = file.split('.')
                    sht.range('B'+str(rows)).value = (rows-8, write_name, file_name, file_type)
                    if file_name.split('_')[0] in filelist_data_dict:
                        sht.range('F'+str(rows)).value = filelist_data_dict[file_name.split('_')[0]]
                    sht.range('K'+str(rows)).value = rule_name%(re.match(match_name, file).group())
                    # row_cell = sht.range('A'+str(rows)).expand('right')
                    sht.range('C'+str(rows)+':F'+str(rows)).color = 0,255,255
                    row_cell = sht.range(str(rows)+':'+str(rows))
                    row_cell.api.Borders(11).LineStyle = 1
                    row_cell.api.Borders(11).Weight = 2
                    row_cell.api.Borders(9).LineStyle = 1
                    row_cell.api.Borders(9).Weight = 2   
                    # sht.range('D'+str(rows)).value = file_name
                    # sht.range('E'+str(rows)).value = file_type
                    rows += 1
                    i += 1
            except Exception as e:
                print('\033[1;31mrow:', rows, '\tfile:', file, '\033[0m')
                print(e)
        if i>2:
            sht.range('K'+str(rows-1)).value = rule_name%(re.sub('\\\\d.*[}+]', 'n', match_name))
            sht.range('K'+str(rows-2)).value = '......'    
            sht.range('K'+str(rows-2)).api.Font.Name = 'Times New Roman'
    sht.range(str(rows-1)+':'+str(rows-1)).api.Borders(9).Weight = 3 
    sht.range('C3').value = rows-1
    wb.save(os.path.join(path, new_name_list[1])) 
    wb.close()
    # wb.app.quit()
except Exception as e:
    print("\033[1;31m生成FileList文件失败\033[0m")
    print(e)
else:
    print("生成FileList文件完成")

print("\n修改主文件中...")
try:
    wb = app.books.open(os.path.join(path, new_name_list[0]))
    print("读取简表中...")
    sht = wb.sheets['简表']
    jianbiao_data_list = sht.range('D8:T14').value
    jianbiao_data_dict = {str(int(i[3])):i for i in jianbiao_data_list}

    data_index = [number]
    if zhuzhang_number not in data_index:
            data_index.append(zhuzhang_number)
    for i in jianbiao_data_dict:
        if i not in data_index:
            data_index.append(i)

    print("修改简表中...")
    sht.range('D8').value = [jianbiao_data_dict[i][:11] for i in data_index]
    sht.range('T8').value = [[jianbiao_data_dict[i][-1]] for i in data_index]

    print("修改封面中...")
    sht = wb.sheets[0]
    sht.range('D15:J17').value = [[jianbiao_data_dict[number][0]], [jianbiao_data_dict[number][9]], [jianbiao_data_dict[number][10]]]
    
    wb.save()
    wb.close()
except Exception as e:
    print("\033[1;31m修改主文件失败\033[0m")
    print(e)
else:
    print("修改主文件完成")

app.quit()
input('\n\n按任意键退出...')