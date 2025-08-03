## 样式设置规则

设置跟随上一行的样式。

`set.follow_styles()`如果是新行，完整复制上一行样式。

`set.follow_styles()`如果是已有行，只在有值的位置复制上一行对应单元格样式。

`set.new_row_styles()`和`r.set.follow_styles()`同时设置时，以后设置的为准。

## 表头格式

xlsx 和 csv 格式支持设置表头格式：

`{None: Header, 'table1': Header1, 'table2': Header2, ...}`

`None`在 xlsx 格式时表示活动数据表，csv 格式则只有`None`一项。

### `Header`

`Header`格式为`{列序号: 表头值}`

`Header[*]`可获取对应列内容，当`*`是`int`时获取该列表头项，当`*`是`str`时获取该表头值所在列序号。

## 数据格式

`Recorder`数据格式：

```python
{None: [],
 'table1': []}
```

`DBRecorder`和`ByteRecorder`的数据保存在`list`
