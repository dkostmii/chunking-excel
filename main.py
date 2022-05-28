import pandas as pd
from pandas import DataFrame
from math import floor

from os import path
from os import mkdir
from shutil import rmtree

from numpy import array_split

outDir = './output'

dataDir = './data'
dataSrc = path.join(dataDir, 'example.xls')


def yn_prompt() -> bool:
  ans = input()
  return ans.lower() == 'y'


def prelude() -> None:
  if not path.exists(dataSrc):
    raise Exception(f"File {dataSrc} does not exists")

  if not path.exists(outDir):
    mkdir(outDir)
  else:
    print("Do you want overwrite existing ./output/ content? [y/N]: ")
    if not yn_prompt():
      exit(0)
    else:
      rmtree(outDir)
      mkdir(outDir)

  print("Length of sheet: ", len(read_excel(dataSrc)))


def read_excel(excel_file: str) -> DataFrame:
  sheet = pd.read_excel(excel_file, index_col=0)
  return sheet


# size - approximate size of each chunk
#
# for size 2000 and length of DataFrame 10000
# it will be equal to 10000 / 2000 = 5
#
# fallbacks to 1
def getChunkCount(df: DataFrame, size: int = 2000) -> int:
  length = len(df)

  size = abs(size)
  length = abs(length)

  if size > length:
    return 1

  return int(floor(length / size))


# uses NumPy to split DataFrame
# as it does with NDArrays
def split_dataframe(df: DataFrame) -> list[DataFrame]:
  chunk_count = getChunkCount(df, size = 2000)

  # Splits DataFrame into list of NDArrays
  nparr_list = array_split(df, chunk_count)

  # Convert NDArrays to DataFrames
  df_list = [
    pd.DataFrame(item, columns = df.columns)
    for item in nparr_list ]

  return df_list


# Gets the first sheet name
# or fallbacks to Sheet 1
def get_first_sheet(excel_file: str) -> str:
  fallback = "Sheet 1"

  xl_file = pd.ExcelFile(excel_file)
  if not isinstance(xl_file.sheet_names, list):
    return fallback

  sheet_name = xl_file.sheet_names[0]

  if not isinstance(sheet_name, str):
    return fallback

  return sheet_name


# Write each dataframe into separate file
# name of files: {originalFile}-{index}.xlsx
def write_dataframes(df_list: list[DataFrame], excel_file: str = dataSrc) -> None:
  sheet_name = get_first_sheet(excel_file=excel_file)
  dataSrcName = path.splitext(path.basename(excel_file))[0]

  for i, df in enumerate(df_list):
    writer = pd.ExcelWriter(path.join(outDir, f"{dataSrcName}-{i}.xlsx"), engine="xlsxwriter")

    df.to_excel(writer, sheet_name=sheet_name)

    writer.close()


def finish(df_list: list[DataFrame]):
  count = len(df_list)
  print(f"Written {count} files")
  print("Done.")


def main():
  # prelude XD
  prelude()

  excel_df = read_excel(dataSrc)
  df_list = split_dataframe(excel_df)
  write_dataframes(df_list, dataSrc)

  finish(df_list)


if __name__ == "__main__":
  # prelude()

  # Uncomment prelude() and comment main()
  # to just check the rows count
  main()
