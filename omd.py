from src import outlook_message_downloader as omd,  yaml_op as yop
import variables as var
import time
from datetime import timedelta

start = time.time()
    
if __name__ == '__main__':
    settings = yop.load_yaml_file(var.DEFAULT_SETTINGS_PATH)
    try:
        messages = omd.get_outlook_messages(settings["root_folder"],
                                            settings["folder_name"])
        dataframe = omd.get_message_attributes(messages, settings["min_date"],
                                               settings["max_date"])

        dataframe.to_excel(settings["output_path"], index=False)
        
        end = time.time()
        print(f"\nProgram succesfully executed")
        print(f"\nExecution time: {str(timedelta(seconds=(end-start)))}")
        
    except PermissionError:
        print("File is busy, or you might not have permissions to write")