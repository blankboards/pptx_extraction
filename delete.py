from modules.config import OUTPUT_DIR_2, OUTPUT_DIR
import shutil

def main():
    shutil.rmtree(OUTPUT_DIR_2)
    print(f"{OUTPUT_DIR_2} DELETED SUCCESSFULLY!")

if __name__ == "__main__":
    main()
