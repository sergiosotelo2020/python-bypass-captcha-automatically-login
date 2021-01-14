import sys
import chilkat

csv = chilkat.CkCsv()
csv.put_HasColumnNames(True)

success = csv.LoadFile("T.csv")
if (success != True):
    print(csv.lastErrorText())
    sys.exit()

success = csv.SetCell(0,22,"baguette")

success = csv.SaveFile("V.csv")
if (success != True):
    print(csv.lastErrorText())