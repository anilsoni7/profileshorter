"""
profile wise sorting
"""
try:
    import argparse
    import collections
    import os
    import sys
    import time

    import xlwt
    import pandas as pd
except ImportError:
    import argparse
    import collections
    import os
    import sys
    import time

    import xlwt
    import pandas as pd


class DataManager:
    """
    DataManager class for profile sorting
    """

    def __init__(self, file_name, search_index,
                 position_location, sheet_index=(0, 1)):
        """
        @param file_name -> str or tuple of file names
        @param build_index -> True or False
        """
        if type(file_name) is type(str):
            self._file_name = tuple(file_name, ) # , due to problem while iterating on tuple
        else:
            self._file_name = tuple(file_name)
        self._position_location = list(position_location)
        self._file_content = None
        self._sheet_index = sheet_index
        self._index = collections.defaultdict(set)
        self._search_index = search_index
        self._student_data = collections.namedtuple('student', ['enrollment', 'index', 'files'])

    @staticmethod
    def genrate_directory_names(_file_name):
        """
        This is static method & only applies for our college problem
        implement for genral use
        genrate directory names based on the file given input
        """
        print("Genrating directory names ")
        for index, name in enumerate(_file_name):
            name = name.split("/")[-1]
            name = name.replace("Interested in ", "")
            name = name.replace(".xlsx", "")
            _file_name[index] = name

    @staticmethod
    def create_directory(_file_name):
        """
        create directory under current directory
        to store data
        """
        print("Creating diretory if not exsist")
        for name in _file_name:
            try:
                os.mkdir(name)
            except FileExistsError:
                pass

    def write_to_file(self, sheet_index, student_index, row, file_index, file_obj):
        """
        write a single row to file
        1 at a time
        eg. Academic with sheet_index = 0
            placement / Training with sheet_index = 1
        """
        for col_index, data in \
                        enumerate(self._file_content[self._file_name[file_index] + \
                                                     str(sheet_index)].loc(0)[student_index]):
            file_obj.write(row, col_index, str(data))

    def read_files(self):
        """
        Read file content and store them in memory
        """
        if self._file_content is None:
            self._file_content = dict()
        try:
            print(self._file_name)
            for _filename in self._file_name:
                print(_filename)
                for _sheet_index in self._sheet_index:
                    print(f"Reading {_filename}  sheet index = {_sheet_index}")
                    self._file_content.__setitem__(_filename + str(_sheet_index),
                                                   pd.read_excel(_filename,
                                                                 sheet_name=_sheet_index)
                                                  )
        except KeyboardInterrupt:
            print("by user")
        except FileNotFoundError:
            assert self._file_name
            print("file not found {}".format(self._file_name))
            sys.exit()

    def build_index(self):
        """
        Content of file in memory and given index name to object (position_location)
        build index based on that
        """
        try:
            # we know all Placement/Training details are on odd index[1,3....] so we will
            # iterate over odd index for building index
            p_i = 0
            for _filename, _content in list(self._file_content.items())[1::len(self._sheet_index)]:
                print(f"Building index for {_filename}")
                position_index = self._position_location.pop(0)

                for index, data in enumerate(_content.values):

                    for post in data[position_index].split(";"):
                        # data[position_index] is the profiles field
                        self._index[post.lower()].add(self._student_data(enrollment=data[1],
                                                                         index=index,
                                                                         files=p_i)
                                                     )#data[1] consist of enrolment
                p_i += 1
        except KeyboardInterrupt:
            print("by user")

    def write_files(self):
        """
        Use the index and store that according to category
        in different files
        """
        #localize the filename for updating
        _file_name = list(self._file_name)
        self.__class__.genrate_directory_names(_file_name)
        self.__class__.create_directory(_file_name)

        saved_training, saved_placement = 0, 0
        for xlfilename, students in self._index.items():
            placement = xlwt.Workbook()
            p_sheets = [placement.add_sheet("Acedemic"),
                        placement.add_sheet("Placement")
                       ]
            training = xlwt.Workbook()
            t_sheets = [training.add_sheet("Acedemic"),
                        training.add_sheet("Training")
                       ]
            row_training = 1
            row_placement = 1
            for student in students:
                if student.files == 0:# training
                    self.write_to_file("0", student.index, row_training, student.files, t_sheets[0])

                    self.write_to_file("1", student.index, row_training, student.files, t_sheets[1])

                    row_training += 1
                else: #placement
                    self.write_to_file("0", student.index, row_placement,
                                       student.files, p_sheets[0])

                    self.write_to_file("1", student.index, row_placement,
                                       student.files, p_sheets[1])

                    row_placement += 1

            xlfilename = xlfilename.replace("/", "-")
            xlfilename = xlfilename.replace(" ", "-")

            if row_placement > 1:
                print(f"{xlfilename}.xlsx in {_file_name[1]}{os.sep}")
                placement.save(f"{_file_name[1]}{os.sep}{xlfilename}.xls")
                saved_placement += 1
            if row_training > 1:
                print(f"{xlfilename}.xlsx in {_file_name[0]}{os.sep}")
                training.save(f"{_file_name[0]}{os.sep}{xlfilename}.xls")
                saved_training += 1

        print(f"Total files {saved_placement + saved_training}")
        print(f"Saved {saved_training} files in {_file_name[0]}")
        print(f"Saved {saved_placement} files in {_file_name[1]}")

    def run(self):
        """
        Main method for DataManager class this try to perform all
        given method in step
        """
        start = time.time()
        self.read_files()
        self.build_index()
        self.write_files()
        finish = time.time() - start
        if int(finish) > 3600:
            print(f"""\n Total time taken to complete job {finish/3600} hour
                  {finish / 3600* 60} minutes {finish / 3600*60*60} seconds """)
        elif int(finish) > 60:
            print(f"""\n Total time taken to complete job {finish / 60} minutes
                   {finish / 3600} seconds """)
        else:
            print(f"\n Total time taken to complete job {int(finish)} seconds ")

    def __del__(self):
        del self._file_name
        del self._position_location
        del self._file_content
        del self._sheet_index
        del self._index
        del self._search_index
        del self._student_data


def main():
    """
    @param
        @file_name -> consisit of 1 or more files for parsing
        @search_index -> are the sheet index of excel files
        @position_location -> is for the location of "Position you will like to apply for"
            file wise supply (position you will like to aplly for) index
            7 is the index of pos.. in file 1
            31 is the index of position.... in file 2
    """
    parser = argparse.ArgumentParser()
    parser.add_argument("-a", "--attribute", help="slect attributes")
    parser.add_argument("-f", "--file", type=list,
                        help="file name or directory followed by filename")
    parser.add_argument("-v", "--values", type=list, help="values to seprate  eg 1,2,3")
    args = parser.parse_args()
    args.file = [r"Interested in Training 2019.xlsx",
                 r"Interested in Placement 2019.xlsx"]
    print(args.file)
    time.sleep(2)

    if args.file is None:
        print("""please include all attributes\n
             python3 seprator.py -a <attribute_name> -f <file_name> -v <values>
            """)
        return
    datamanager = DataManager(file_name=list(args.file), search_index=[0, 1],
                              position_location=[7, 31])
    datamanager.run()

    del datamanager

if __name__ == "__main__":
    main()
