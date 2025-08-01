# Flat-OPC-to-OOXML
Convert a Flat OPC file to an Open Office XML package library file type using python

Usage
---
```python
from flat_to_opc import flat_to_opc

flat_to_opc(input, output_directory_path_string)
```
This code either takes a stream input from a file like object like BytesIO or StringIO or a string containing the path to the the Flat OPC file and converts the contents into the Open Office XML equivalent before writing it to the the file path specified in string format.
<br/>
<br/>
<!-- -->
```python
from flat_to_opc import flat_to_opc_bytes

output = flat_to_opc_bytes(input)
```
Same as the previous function usage however instead of writing to a file location, returns the Open Office XML as binary data from a BytesIO buffer.
