# **<center>EasyWord 2</center>**


## EasyWord 2是什么
​	2.x版本是对于1.x的重构和升级

## EasyWord2如何使用

```
replaceLabel(@NotNull InputStream inputStream,
                                    @NotNull OutputStream outputStream,
                                    @NotNull Map<String, Customization> label)
```

​		2.x版本将实现抽离，现在不再区分标签类型。换言之，easyword2.x的核心代码变成了一个标签定位器，在定位到标签后将标签所在位置的所有元素暴露出来，我们的标签内容实际上就是实现。

1.x标签简介：

	静态标签：在word的任意位置插眼（因为EasyWord是基于run实现的标签匹配，所以打标签要注意在一个run内），用于进行占位替换操作的标签；
	动态标签（仅限paragraph）：在word的paragraph中的标签，回填会一行一行的回填；
	列表标签（仅限table中）：在word的table中的标签，回填的时候会按照table的行一行一行的回填；
	图片标签：其属性和静态标签一致，唯一的区别是进行图片回填；
	列标签（仅限table中）：在word的table中的标签，回填的时候会按照table的列一列一列的回填；
	隐藏标签：以上标签均支持隐藏标签，实际上隐藏标签并不是指一种标签类型，而是将打的标签隐藏起来从而在看模板的时候看不到标签；

在统一标签后，实现了1.x的几种标签：

#### 1. 静态标签的替换

​		现在由  StaticLabelImp.class实现

#### 2. 动态标签的替换

​		现在由  DynamicLabelImp.class实现

#### 3. 列表标签的替换

​		现在由  TabelLabelImp.class实现

#### 4. 图片标签的替换

​		现在由  PictureLabelImp.class实现

#### 5.列标签的替换

​		现在由  VerticalLabelImp.class实现

#### 6.隐藏标签的替换
​		隐藏标签在代码部分并没有什么不一样，区别在于往模板中打的标签是否是隐藏状态（关于word如何隐藏文字请自行百度）。无论标签是否是隐藏状态都会被检测到，而且一旦替换成功，EasyWord会将替换后的内容由隐藏状态变为可见状态。

#### 7.Word合并
```java
EasyWord.mergeWord(wordList, outputStream);
```
wordList：需要合并的文件流集合，按先后顺序进行合并；  
outputStream:合并后文件的输出流;  
word合并后每两个word之间会默认加一个换页符，目前没有开发定制化接口进行开关。  

#### 8.高级

​		上述常用标签的默认实现若是不能满足开发者需求，可自行实现Customization接口

## Word如何打标签
​		不同于easyword1.x标签必须在一个XWPFRun里面，2.x版本将标签位置扩展到整个paragraph，以后打标签更加方便啦

## 版本

**2.0.0** 2019-15-6

* 重构easyword发布
