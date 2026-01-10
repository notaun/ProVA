# LISTING ITEMS
# listing everything in the given folder
# (. = current directory)
import os
import shutil

def list_items(path="."):
    try:
        items = os.listdir(path)
        if items:
            print("\nItems in directory:")
            for item in items:
                print(" -", item)
        else:
            print("\n ️ ⚠️ ️Directory is empty.")
    except Exception as e:
        print(f"❌ Error listing items: {e}")


# CREATE NEW FOLDER
# making a new folder with the given name, if folder already exists, it shows a warning instead of crashing
def create_folder(name):
    try:
        if not os.path.exists(name):
            os.mkdir(name)
            print(f"Folder '{name}' created.")
        else:
            print(" ⚠️ Folder already exists.")
    except Exception as e:
        print(f"❌ Error creating folder: {e}")


# DELETE FOLDER
# deleting a folder safely (but only if exists and is actually is a folder)
def delete_folder(name):
    try:
        if os.path.exists(name) and os.path.isdir(name):
            confirm = input(f"⚠️ Are you sure you want to delete the folder '{name}'? (yes/no): ").strip().lower()
            if confirm in ["yes", "y"]:
                shutil.rmtree(name)
                print(f"Folder '{name}' deleted successfully.")
            else:
                print("❎ Deletion cancelled.")
        else:
            print(f"❌ Folder '{name}' not found.")
    except Exception as e:
        print(f"❌ Error deleting folder: {e}")
# here i used shutil.rmtree() so it can remove folders even if they contain files


#CREATE NEW FILE
#creating a file
def create_file(name):
    try:
        if not os.path.exists(name):
            with open(name, 'w') as f:
                f.write("")  # create empty file
            print(f"File '{name}' created.")
        else:
            print("⚠️ File already exists.")
    except Exception as e:
        print(f"❌ Error creating file: {e}")


# DELETE FILE
# deleting files safely (checks existence first)

def delete_file(name):
    try:
        if os.path.exists(name) and os.path.isfile(name):
            confirm = input(f"⚠️ Are you sure you want to delete the file '{name}'? (yes/no): ").strip().lower()
            if confirm in ["yes", "y"]:
                os.remove(name)
                print(f"File '{name}' deleted successfully.")
            else:
                print("❎ Deletion cancelled.")
        else:
            print(f"❌ File '{name}' not found.")
    except Exception as e:
        print(f"❌ Error deleting file: {e}")


# RENAME FILE/FOLDER
#rename (or moves) a file/folder from one name to another
def rename_item(old, new):
    try:
        if os.path.exists(old):
            os.rename(old, new)
            print(f"Renamed '{old}' → '{new}'.")
        else:
            print(f"❌ Item not found")
    except Exception as e:
        print(f"❌ Error removing item: {e}")

if __name__ == "__main__":
    print("\n--- TESTING FILE MANAGER ---\n")

    list_items()
    create_folder("TestFolder")
    create_file("testfile.txt")
    rename_item("testfile.txt", "renamed.txt")
    delete_file("renamed.txt")
    delete_folder("TestFolder")

    print("\n--- FINAL CHECK ---")
    list_items()
