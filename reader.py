class ReadFile:

    #
    def __init__(self, filetype):
        self.filetype = filetype

    #


readfile = ReadFile("file.csv")
print(readfile.filetype)