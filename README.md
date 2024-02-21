# create_print_onboard_word_documents
根据表格数据，生成多张入职单，并合并成一个进行打印
### 需求背景

每当一个企业人比较多了之后，每周就会有大量的人入职离职。入职的时候需要把入职信息，主要是一些平台的账号密码，和一些说明发给新员工。

手动输入肯定是不可能的，动不动就入职十几个人。

我们一开始是用word的批量打印功能，从excel导入数据然后批量打印那个，但感觉那个操作起来也有点烦，所以想着用python写个自动生成入职单。

### 工作步骤分解

1. 在Excel文件中录入入职人员信息
2. 读取Excel文件
3. 读取Word模板文件
4. 按人数做循环：替换占位内容
5. 保存至新文件，生成N个文档
6. 合并文档到一个文档，方便打印
### 1&2.从Excel读取入职信息

数据位置对应如下：
| 科目 | 键值 | 所在列 | 示例值 |
| --- | --- | --- | --- |
| 工号 | id | A | 06479 |
| 姓名 | name_cn | B | 刘起 |
| 英文名 | name_en | C | Kelly Liu |
| 部门 | dep | E | RM |
| 分机号 | tel | I | 021-5229-8706 |
| 计算机名 | PC | J | CAP-2020 |
| 邮箱 | ad_account | K | kel.liu |
| 邮箱密码 | password | L | abc@1234 |

```python
def read_from_excel():
    """
    从入职信息excel读取信息：resources/入职信息.xlsx
    :return: 返回得到的员工信息列表
    """
```
### 345.获取word模板，生成n个入职单

```python
def read_word_temp():
    """
    读取word模板
    :return: 返回Document对象
    """
    doc = Document("resources/入职信息模板.docx")
    return doc
def make_onboard_words(employees):
    """
    根据员工数量生成n张入职单
    :param employees: 员工信息的列表
    :return: 成功返回True
    """
```
### 6.合并到一个文件

用到这个包：

```python
pip install docxcompose
```

最好保证文档刚好一页，这样就不需要加空白页

```python
def get_files_list(source_path):
    """
    获取指定文件路径的文件名列表
    :param source_path: 文件夹路径
    :return: 文件名列表
    """
    
    
def make_words_one(source_files, target_file):
    """
    将多个word合并成一个
    :param source_files: 源文件列表
    :param target_file: 合并的目标文件路径+名称
    :return: 成功返回True
    """
```
主程序：调用以上方法，实现最终目的

```python
def onboard_main():
    # 从入职信息excel读取信息
    employees = read_from_excel()
    print(f"获取到{len(employees)}条入职信息--->")
    # 根据入职信息生成N份入职表单
    if make_onboard_words(employees):
        print("生成入职单成功--->")
    source_path = "onboard_words/"  # 源文件路径
    target_file = "resources/入职单打印.docx"  # 目标路径
    # 获取源文件夹的文件列表，合并为一个新文档
    make_words_one(get_files_list(source_path), target_file)
    print(f"合并成功---^_^")
```
### How to use

您可以直接在Pycharm中打开运行，这样方便修改参数；
也可以运行exe文件，接下来说明下如何使用

程序同目录下有两个文件夹：

- resources用来放入职信息.xlsx、入职信息模板.docx
- on_board_words用来放生成的多个入职word文件

首先，在入职信息里录入入职的人的信息，表格结构请参考：resources/入职信息.xlsx


然后，运行主程序create_onboard_files.exe即可

生成的多个入职单，保存在on_board_words

生成的 入职单打印.docx，保存在resources里，需要打印时，打印该文件即可。

### 生成签名

在excel中填写信息，然后生成word文档用作签名。

运行以下路径的excel文件，然后运行python程序即可

/**gene_sign**/**[gene_sign.py](https://github.com/h-kayotin/kayotin_doc_excel/blob/master/gene_sign/gene_sign.py)**
