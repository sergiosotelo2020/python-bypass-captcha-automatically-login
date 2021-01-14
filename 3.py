import sys
import chilkat

csvv = chilkat.CkCsv()
csvv.put_HasColumnNames(True)

success = csvv.LoadFile("T.csv")
if (success != True):
    print(csvv.lastErrorText())
    sys.exit()

success = csvv.SetCell(0,22,"baguette")

success = csvv.SaveFile("V1.csv")
if (success != True):
    print(csvv.lastErrorText())