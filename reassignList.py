import xlrd
import xlwt
import math

def createRemovedList(databaseChanges):
  """
  Reads the changes column of the database changes output and creates a list of reps that need to be reassigned
  back to their accounts
  """
  # read workbook
  file = databaseChanges
  workbook = xlrd.open_workbook(file)
  #open workbook
  sheet = workbook.sheet_by_index(0)

  # write to workbook
  wbt = xlwt.Workbook()
  wst = wbt.add_sheet('Reassignments')

  endOfFile = 2180
  entryIndex = 0

  for x in range(2, endOfFile):
    stringToTest = sheet.cell_value(x,3)
    if "was removed" in stringToTest:
      entryIndex += 1
      stringToTest = stringToTest.split("rep",1)[1]
      repName = stringToTest.split("was",1)[0]
      stringToTest = stringToTest.split("account", 1)[1]
      accountName = stringToTest.split(".", 1)[0]
      wst.write(entryIndex, 0, repName)
      wst.write(entryIndex, 1, accountName)

  #save finished workbook
  newListName = "ReassignmentList.xls"
  wbt.save("ReassignmentList.xls")
  print("file saved as " + newListName)


databaseChanges = input("What is the name of the file (without extension)?")

if databaseChanges == "":
  databaseChanges = "DatabaseChanges"

databaseChanges = databaseChanges + ".xlsx"

createRemovedList(databaseChanges)