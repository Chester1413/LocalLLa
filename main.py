'''from newtool import *   導入全部'''
from newtool import save_conversation_to_docx
from newtool import get_user_input
from newtool import send_request_to_oobabooga

def main():
    history = []
    first_time = True
    while True:
        folder_path, prompt = get_user_input(first_time)
        first_time = False
        if prompt.lower() == "exit":
            print("退出對話。")
            break
        elif prompt.lower() == "exit2":
            save_conversation_to_docx(history)
            print("退出並保存對話。")
            break
        send_request_to_oobabooga(folder_path, prompt, history)

if __name__ == "__main__":
    main()