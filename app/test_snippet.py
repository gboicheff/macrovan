from utils import *


#iterate through folder_dict and create a subfolder copying the files over for each organizer
#def create_folders(folder_dict, parent_folder_name):
if __name__ == '__main__':
    parent_path = r'D:\Stuff\Projects\Pol\macrovan\io\Output'
    print('parent_path', parent_path)
    os.chdir(parent_path)
    print('os.getcwd() = ', os.getcwd())
    parent_folder_name = 'Test Folder'
    if(os.path.isdir(parent_folder_name)):
        print('Found Parent Folder, about to remove')
        shutil.rmtree(parent_folder_name)
    os.mkdir(parent_folder_name)
    os.chdir(parent_folder_name)
    print('parent_folder_name = ', parent_folder_name)
    #print('folder_dict keys = ', folder_dict.keys())
    #for subfolder in folder_dict:
    # for key in folder_dict.keys():
    #     if key == '':
    #         key = 'no_org'
    #         folder_dict['no_org'] = 'no_org'
    #     subfolder = key
    #     print('subfolder = ', subfolder)
    #     os.mkdir(subfolder)
    #     os.chdir(subfolder)
    #     for file in folder_dict[subfolder]:
    #         search_file = file + "*" + ".pdf"
    #         search_file = search_file.replace(" ", "")
    #         #for file in os.listdir(parent_path+"\io\output"):
    #         for file in os.listdir(parent_path+r"\io\Output\tests"):
    #             found_file = file.replace(" ", "")
    #             print('search_file = ', search_file)
    #             print('found_file = ', found_file)
    #             print()
    #             if fnmatch.fnmatch(found_file, search_file):
    #                 shutil.copy(parent_path+r"\app\io\output\\tests"+file, file)
    #                 break
    #     os.chdir("..")
    # os.chdir(parent_path)