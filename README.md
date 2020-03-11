# MOVIE_DATA
将2010-2017年电影的一些数据进行收集整理。

&nbsp;
&nbsp;

### 1、上映电影以及票房数据的获取

利用boxoffice文件夹里的程序，将58921网站上2010到2017年的电影票房数据爬取下来，存放地址为data/boxoffice。  

&nbsp;
&nbsp;
 
### 2、从豆瓣上爬取电影的相关信息

利用douban文件夹里的程序movie_cralwer_with_api，先将与电影有关的豆瓣页面网址爬取下来（利用了豆瓣的api接口，进行电影片名的的搜索，找出片名类型为movie的数据，先根据电影片名进行筛选，选取正确的网址，再根据年份进行筛选，选取正确的网址，对于还是未能找到相关网页的数据，进行人工处理，添加或删去）数据存放地址为data/douban。通过豆瓣提供的api，将电影的相关信息读取。
&nbsp;
利用douban文件夹里的程序douban_movie_crawler，将豆瓣页面上的信息爬取下来，信息包括电影的片名，编剧，类型，导演，想看人数，主演，时长，上映日期。同时对电影的信息进行校对，针对上映日期错误的，进行人工的校正或剔除，对于因为电影名相同而错误爬取，进行人工的校正或剔除，对于爬取错误信息的电影，进行人工的校正或剔除。最终将数据保存在data/movie里。
&nbsp;
利用exact文件夹里的程序，对爬下来的数据进行整理，并写到data/exact中的表格里。针对表格里空白的数据，进行人工填充。
&nbsp;
&nbsp;
 
