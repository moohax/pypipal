# pypipal
Simple script that implements Pipal in python. Thanks to the original authors of pipal and @culturedphish for providing the excel stuff.

## Requirements
Uses Python3 and Pandas
```python
pip install pandas
```
A file of hashes should look as follows:

```
<hash>:<cracked_hash>
<hash>:
<hash>:<cracked_hash>
<hash>:
<hash>:
```

## Examples:
```python
pypipal.py -f hashes.csv -o analysis.xlsx
```

## To Do:
- Move away from Pandas
