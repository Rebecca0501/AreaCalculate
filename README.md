# AreaCalculate

The purpose of this program is to assist architects in quickly calculating the total area of different zones on each floor.

After the program executes, it will generate:

1. **An Excel file** containing records of "floor area", "hall area", "management committee area", "mechanical and electrical area", "driveway area", and "balcony area" for each floor.
2. **A DXF file** named "original filename-after.dxf", which adds textual descriptions of the area review onto the original DXF file.
> [!NOTE]
> Note: This project primarily assists in drafting architectural license drawings in Taiwan, so the textual descriptions of the area review are mainly based on Taiwan's building regulations.

這個程式項目是為了協助建築設計師 "快速" 計算每個樓層不同區域的總面積

程式執行完畢後，會生成
1. **Excel檔**，裡面記錄了各個樓層的  "樓地板面積"、"梯廳面積"、"管委會面積"、"機電面積"、"車道面積"、"陽台面積"
2. **DXF檔**，名稱為"原始檔名-after.dxf"，此檔案在原始的dxf檔上加上面積檢討的文字敘述
> [!NOTE]
> 本專案主要是協助繪製台灣的建築執照圖，因此面積檢討的文字主要是根據台灣的建築法規規定闡述)


## Installation

Before running the project, you must use the package manager [pip](https://pip.pypa.io/en/stable/) to install openpyxl and ezdxf.

```bash
pip install openpyxl
pip install ezdxf
```

## Usage

在Step1. Get CAD file path區域修改代碼以找到需要處理的CAD檔
> [!NOTE]
> 程式只能處裡副檔名為dxf的檔案
```python
#Step1. Get CAD file path
str_CAD_file_name="example.dxf"
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License

[MIT](https://choosealicense.com/licenses/mit/)
