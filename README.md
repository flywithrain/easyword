# **<center>EasyWord</center>**


## EasyWord是什么
​	EasyWord是基于POI实现的对docx格式文件进行各种操作的工具。
​	作者在好几个项目中遇到对word的操作，这些代码基本都是重复模板代码，加上写代码的人对POI不熟悉或者POI版本自身bug导致或多或少的一些问题，为了以后偷懒做出这么个东西。

## 在哪里下载EasyWord
github：https://github.com/flywithrain/easyword
EasyWord毕竟难登大雅之堂，所以中心仓库是没有的，大家要用可以放到自己本地仓库或者私有仓库

## EasyWord如何使用
	静态标签：在word的任意位置插眼（因为EasyWord是基于run实现的标签匹配，所以打标签要注意在一个run内），用于进行占位替换操作的标签；
	动态标签（仅限paragraph）：在word的paragraph中的标签，回填会一行一行的回填；
	列表标签（仅限table中）：在word的table中的标签，回填的时候会按照table的行一行一行的回填；
	图片标签：其属性和静态标签一致，唯一的区别是进行图片回填；
	隐藏标签：以上四种标签均支持隐藏标签，实际上隐藏标签并不是指一种标签类型，而是将打的标签隐藏起来从而在看模板的时候看不到标签；

EasyWord对外提供的方法都在com.thunisoft.easyword.core.EasyWord类中一共有3类方法，替换replaceLabel和replaceLabelite，合并mergeWord
以下所有方法均可在github上找到示例（EasyWordExample）

#### 1. 静态标签的替换

```java
EasyWord.replaceLabel(fileInputStream, fileOutputStream, staticLabel);
```
<span id="in">fileInputStream:这是模板的输入流；</span>  
<span id="ou">fileOutputStream：这是替换后文件的输出流；</span>  
staticLabel：静态标签，类型是Map<String, Customization>，其中String就是模板中的标签，EasyWord采用replaceAll去替换String标签，替换内容即Customization中文字。  

#### 2. 动态标签的替换

```java
EasyWord.replaceLabel(fileInputStream,
                fileOutputStream),
                new HashMap<>(0),
                dynamicLabel,
                new HashMap<>(0),
                new HashMap<>(0));
```
fileInputStream:<a href="#in">同上</a>;  
fileOutputStream:<a href="ou">同上</a>;  
dynamicLabel:动态标签，类型是Map<String, List\<Customization>>，其中String就是模板中的标签，每一个Customization就是段落中的一行，List\<Customization>就是回填内容的集合

#### 3. 列表标签的替换

```java
EasyWord.replaceLabel(fileInputStream,
                fileOutputStream),
                new HashMap<>(0),
                new HashMap<>(0),
                tableLabel,
                new HashMap<>(0));
```
fileInputStream:<a href="#in">同上</a>;  
fileOutputStream:<a href="ou">同上</a>;  
tableLabel:表格标签，类型是Map<String, List<List\<Customization>>>，其中String就是模板中的标签，和dynamicLabel很相似，每一个Customization就是一个单元格（cell），每一个List\<Customization>就是一行（row）,自然的List<List\<Customization>> 就代表rows组成的表格  

#### 4. 图片标签的替换

```java
EasyWord.replaceLabel(fileInputStream,
                fileOutputStream),
                new HashMap<>(0),
                new HashMap<>(0),
                new HashMap<>(0),
                pictureLabel);
```
fileInputStream:<a href="#in">同上</a>;  
fileOutputStream:<a href="ou">同上</a>;  
pictureLabel:图片标签，类型是Map<String, Customization>，其中String就是模板中的标签，Customization中需实现getPicture（）以及图片相关的方法（当getWidth()和getHeight()方法获取到的像素是非自然数时EasyWord会按照图片原始大小展示）  

#### 5.隐藏标签的替换
隐藏标签在代码部分并没有什么不一样，区别在于往模板中打的标签是否是隐藏状态（关于word如何隐藏文字请自行百度）。无论标签是否是隐藏状态都会被检测到，而且一旦替换成功，EasyWord会将替换后的内容由隐藏状态变为可见状态。

#### 6.Word合并
```java
EasyWord.mergeWord(wordList, outputStream);
```
wordList：需要合并的文件流集合，按先后顺序进行合并；  
outputStream:合并后文件的输出流;  
word合并后每两个word之间会默认加一个换页符，目前没有开发定制化接口进行开关。  

#### 7.高级

EasyWord的DefaultCustomization默认按照模板标签所在的样式进行回填替换，虽然大部分情况下我们都可以通过调整模板的样式来控制回填样式，但是如果有些样式不能通过模板来实现或者说在制作模板阶段压根还不知道样式怎么办呢？这时候我们需要实现Customization接口。  
Customization接口中handle方法能够获取到标签回填时刻替换内容所在的table、row、cell、paragraph、run以及他们对应的Index（关于word的结构比这不再详细赘述，要实现定制化需要对word以及POI有一定了解）

## Word如何打标签
打标签要对word结构有一定了解。标签只有一个要求，即标签文字需要在docx文件中一个<w:r></w:r>内。如何检验标签正确性呢，用解压缩软件打开docx文件查看./word/document.xml文件找到标签文字，看其是否在同一对<w:r></w:r>包裹中。document.xml也是学习docx文件结构一个好的入口。

## 版本

- **Aplha** 2019-08-19
	* 从项目中抽离代码，初始化项目；
	* 增加单元测试，修复部分bug；

- **Beta** 2019-08-23
	* 完善隐藏标签替换功能；
	* 新增word合并功能；
	* 修复所有已知bug；

- **1.0.0** 2019-08-24
	* 新增replaceLabelite方法，简化替换操作
	* 修复Customization设置WordConstruct不正确bug

- **1.0.1** 2019-09-20
  - 修复pom引用文件中类型<type>缺失bug
  - 修复表格回填图片清空cell导致index出错bug