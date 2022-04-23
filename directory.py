import os
def prerequisitedirectories():
    try:
        os.mkdir("./result")
        os.mkdir("./data")
    except:
        print("Folders Present")


if __name__ == "__main__":
    prerequisitedirectories()
