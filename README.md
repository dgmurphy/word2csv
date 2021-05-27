# Word2CSV

The application is designed to be run as a Windows EXE--download it from the "Releases" section, then drag and drop your Word file onto the app file, and the output CSV will be created with the same path as the input file, but with the "docx" extention replaced by "csv".

It's also possible to run the program from the command line:
$ ./word-2-excel.exe <input file path>

# Creating an EXE File
To crate an EXE file for the host OS, run "create_exe.sh" from any terminal that support shell scripts (e.g., on Windows, it's convenient to use the MINGW64 packaged with Git). The new file will be written to the "Dist" directory (and overwrite any existing file--the one included in the repository is for Windows).


