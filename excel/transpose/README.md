# excel 表格处理之数据转置
## 用于将 excel 表格中的某行数据，转置为指定容量的多个平行列

## 例如：
```
A B C D E F G H I J K L M

转置为 (单列容量为4， 间隔为2 的数据)

A  E  I  M
B  F  J
C  G  K
D  H  L
```
## 程序执行后，会保存为新的文件 "new_{旧文件名}"