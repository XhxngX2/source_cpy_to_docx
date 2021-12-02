# source_cpy_to_docx
软件著作需要把源代码拷贝到一个Word中，此项目迭代文件夹将制定文件后缀的文件内容复制到单个word文件中。



## arguments

### dir

源代码所在的文件夹。



### file

源代码的项目类型。由于不同的源代码工程需要关联的源代码文件后缀不同，如**c/c++**工程关联的文件主要有：

`.c`,`.h`,`.cpp`，而**python**工程主要关联`.py`文件。

项目默认支持的项目选项为：`C/C++`, `Python`*（Line:54）*

```python
parser.add_argument('-f', '--file', choices=['C/C++', 'Python'],
                        required=True, help="\n\rSpecify the Project kind.")
```

对应支持的文件后缀列表为如下（Line:4~5）

```
C_CPP_FILE_ENDINGS = [".c", ".h", ".cpp", ".asm"]
PYTHON_FILE_ENDINGS = [".py"]
```

