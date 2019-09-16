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
    implementation 'com.github.hao896259037:android-word2html:1.0.0'
}
```
### Example
```Java
WordUtils.getInstance().word2html(
                sourceFilePath,//word file path
                tempDir,//after conver html file storage path 
                htmlFileName);//html fileName
```
