About all: Las files are very common in O&G sphere, and here is some scripts. I find them pretty useful for fast analysis, or sometimes for manipulating the files, and loading them to Petrel.

============================================

Extract_mnemonics:

Input: A folder with Las files.

Output:.xlsx file with mnemonics Data in Las files.

Comments: The file "!_mnemonics.xlsx" will be created in folder with lases, containing two sheets, one - with links to files, and mnemonics inside them, and second one, with each mnemonics, the number of files it is contained in, and the description to it.
============================================

Get_las_info

Input: A folder with Las files.

Output: A .xlsx file, with the Data about the las files

Comments: An output file will contain wellname, bitsize, company, interpretation company, etc.
============================================

Make_PointWellDataFile

Input: A folder with folders, with the names of objects(BedBoundaries, fracture, etc), with the las files inside them.

Output: A .xlsx file, easily transformable by hands into PointWellDataFile.

Comments: You need to first sort your UBI, FMI files by objects, and them run the script on that folder. It will make a giant DataFrame, and write it to Excel.
============================================


Sort_lases_by_wells
Input: A folder with las files, the folder where you want to have sorted files

Output: folders named after the wellnames, and containing the files inside them.
============================================

MergeLasFiles
Input: A folder with LAS files, containing the "[NameOfTheWellInYourPetrelProject]_" before the actual filename (may take some work to detect, which file corresponds to which well).

Output: Folder with folders named after wells, with the files inside them, and the !_Result folder, containing the summary data for each well

Comments: This is a big one. So, generally, you will have several Las files for a well, they can have different curves and MDs. Petrel (at least 2021), doesn't let you to load them easily (it can either overwrite the files entirely, or add the Data to the existing one, but they must not overlap). So my solution was to make one load file from them all, containing all the curves and all data inside them.

The script is run with the main.py file, i recommend to try it on 2-3 wells, and when you understand how it works, go on.
