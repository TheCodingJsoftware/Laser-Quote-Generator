# Laser-Quote-Generator

Generate quotes at the speed of light :)

## Requirements

Create a virtual environement

```
virtualenv venv
```

Activate it and run:

```
pip install -r requirements.txt
```

## Build

**EDIT THESE LINES IN THE .spec FILE**

`a.datas += Tree("C:/Users/jared/AppData/Local/Programs/Python/Python39/Lib/site-packages/grapheme/", prefix= "grapheme")`

`a.datas += Tree("C:/Users/jared/AppData/Local/Programs/Python/Python39/lib/site-packages/about_time/", prefix= "about-time")`

**MAKE SURE YOUR** `C:/Users/jared/AppData/Local/Programs/Python/Python39` **PATH IS CORRECT**


Then install with:

```
pyinstaller main.spec
```

## Demo

![image](https://user-images.githubusercontent.com/25397800/173172193-9c50fe04-9370-464f-9c94-c9c224893cda.png)
