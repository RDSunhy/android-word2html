# android-word2html

### Add Library
```Java
allprojects {
    repositories {
        ...
	maven { url 'https://jitpack.io' }
    }
}

dependencies {
    implementation 'com.github.hao896259037:android-word2html:1.1.0'
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
WordUtils.getInstance(basicSet).word2html();
```

### Known Bug
```Java

  I can't set the image to fit the screen width
  I tried to style the label
  //String imgBegin = "<img src=\"%s\" width=\"100%\" height=\"auto\">";
  //basicSet.setImgBegin(imgBegin);
  But it failed...

```