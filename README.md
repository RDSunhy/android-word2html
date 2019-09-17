# android-word2html

### CSDN博文  [android word文档预览(支持doc/docx两种格式)](https://blog.csdn.net/qq_38356174/article/details/100915969)

### Add Library
```Java
//build.gradle
allprojects {
    repositories {
        ...
	    maven { url 'https://jitpack.io' }
    }
}
//app/build.gradle
android{
    defaultConfig{
        ...
        multiDexEnabled true
    }

    compileOptions {
        sourceCompatibility JavaVersion.VERSION_1_8
        targetCompatibility JavaVersion.VERSION_1_8
    }
}

dependencies {
    implementation 'com.github.hao896259037:android-word2html:1.2.0'
}


//If this fails
//You can Import Module (android_word2html)

```
### Example
```Java
BasicSet basicSet = new BasicSet(
        MainActivity.this,
        sourceFilePath,//word file path
        htmlFilePath,//after conver html file storage path
        htmlFileName);//html fileName
//Some configuration can be added...
//The concrete in BasicSet.class
//basicSet.setHtmlBegin(htmlBegin);
String htmlSavePath = WordUtils.getInstance(basicSet).word2html();

...
//Render in a webview
webView.loadUrl("file://"+htmlSavePath);

```

### Known Bug
```Java

  I can't set the image to fit the screen width
  I tried to style the label
  //String imgBegin = "<img src=\"%s\" width=\"100%\" height=\"auto\">";
  //basicSet.setImgBegin(imgBegin);
  But it failed...

```
