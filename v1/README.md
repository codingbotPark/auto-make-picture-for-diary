# auto-make-picture-for-diary
다이어리에 붙일 사진을 ppt에 하나하나 넣는 과정이 귀찮아서 만들기 시작

아는사람은 알지만 나는 아날로그 다이어리를 쓴다, 다이어리는 쓰면 다 좋다고 생각할거지만 좋은점은 하루로 보면 하루를 돌아볼 수 있고, 예전을 보면 예전에 내가 했던 것들을 보며 **특히 그 당시의 나의 생각과 느낌**을 생각할 수 있어서 좋은 것 같다,

굳이 아날로그를 쓰는 이유는 **사실 그냥 일기장을 산거긴 하지만*, 아날로그로 다이어를 썻을 때 더 프라이빗하고, 애정이 가는 느낌이다, 무엇보다 개발만하다보면 노트북만 계속하게 되는데 한 번씩 리프레시가 가능하다

좋은점은 더 말할 수 있지만 그만하겠다, **암튼! 나의 일기장에 조금 문제가 생겼다** 작년부터 써왔지만 요즘 초심을 잃어서 사진을 뽑아 붙이는게 귀찮아 졌다, **그렇게 하루하루 지나고 보니 반년이 지나서 이젠 손으로 못할 수준이 됐다**

![일기 사진](https://user-images.githubusercontent.com/85085375/178497572-45c3eef1-c2d0-437e-8197-18f9522de5d1.png)

엄 그래서 만들게 됐다, 그래도 한 번 만들면 앞으로 내 인생의 시간을 단축해줄거라 생각한다 드가자


## 1. `python-pptx` 설치(220712)
```
pip install python-pptx
```

pptx에서 파이썬의 ppt제어를 돕는다고 한다, 기본적으로 아래와 같이 설정했다

```py
from pptx import Presentation
from pptx.util import Inches

# ppt 열기
prs = Presentation()
# ppt 슬라이드 추가
slide = prs.slides.add_slide(prs.slide_layouts[6]) # 슬라이드 추가

# ppt저장 
prs.save('다이어리사진.pptx')
```

## 2. ppt에 넣을 사진의 경로 읽기
특정 경로의 폴더안에 있는 사진들을 다 ppt에 넣는 방식을 사용할 것이기 때문에 경로를 지정하고, 지정한 경로의 파일들을 읽는 코드를 작성한다

```py
print("사진이 있는 폴더 경로를 입력하세요 : ",end=" ")
path = input()
fileList = os.listdir(path)
```

`os` 라는 라이브러리를 활용해 폴더와 파일을 다룰 수 있다, 나는 os에 폴더 위치를 주면 안의 파일들을 알아내는 용도로 사용했다

### 3. ppt에 사진을 추가한다
위에서 추출한 파일들을 ppt에 삽입한다, 삽입 메소드는 `슬라이드.shapes.add_picture()` 함수를 사용해 넣을 수 있다

```py
# 매개변수는 각각 파일, left, top, width, height 이다
pic = slide.shapes.add_picture(path+"\\"+name,0,0,Inches(1.3)) # 3.3
```

> 명령어를 보면 width만 줬는데, width만 주면 height는 비율에 맞게 정해진다

위처럼 slide에 사진을 추가할 수 있다
요기까지는 구글링 했을 때 자료가 많아서 큰 문제없이 할 수 있었다

하지만 **이 땐 몰랐다 Inche를 사용하는게 나를 앞으로 어떻게 괴롭힐지!**, 분명 인치 말고 다른 방법이 있을거같다 내일 찾아보겠다

### 4. 사진 돌리기
ppt에 넣은 사진들을 프린팅해서 다이어리에 붙여야 하는데, 종이의 낭비를 줄이고 사진의 일관된 크기를 위해 비율에 따라 사진을 돌린다

위의 사진 추가 명령어를 보면 `width`와 `height` 를 메소드를 정하는데, **만약 세로가 더 긴 사진이라면90도 돌려줘야한다**

```py
width,height = pic.image.size # 사진 사이즈를 각각 width, height에 넣음

if (height > width):
    pic.rotation = 90 # 돌리기
    # 가로 세로를 변경해줘야 다음 위치를 정하기 쉽다
    width,height = height,width
```
> `rotation` 이라는 속성을 수정해서 쉽게 돌릴 수 있었다, 앞으로 더 계산이 필요할걸 생각해서 width와 height를 변경해줬다

### 5. 사진들의 사이즈 조절
내가 구글링을 잘 못해서 그런건진 몰라도 `python-pptx` 라이브러리 한국어 자료가 많이 없었던 것 같다, 그래서 "사진 사이즈 조절" 정도의 깊이만 들어가도 영문으로 검색해서 힌트를 얻을 수 있었다

그래도 `python-pptx` 가 꽤 직관적인 것 같아서 구글링에 없었던 내용이였는데, 아래처럼 사이즈를 조절할 수 있었다, 이게 됐을 때 뿌듯했다ㅋㅋㅋ

```py
pic = slide.shapes.add_picture(path+"\\"+name,0,0,Inches(1.3)) # 3.3
...
pic.height = Inches(1.3)
pic.width = round((Inches(1.3) * (width / height)))
```

아무튼 내가 해야됐던 것은 **평소 다이어리에 넣는 사이즈대로 3.3cm로 맞춰야 하는데, 정방형(3.3 X 3.3), 3대4(3이 3.3), 16대9(9가 3.3)의 사이즈로 맞춰야 했다** 그래서 위에서 돌린 요소는 가로 3.3 세로 3.3 으로 사진이 망가지기 때문에 다시 원래대로 돌리는 작업이다

![결과](https://user-images.githubusercontent.com/85085375/178494150-9407241e-2968-4380-86a3-e95bb405f0cc.png)

> 위처럼 이미지가 생기고, 사이즈 비율이 다 맞게 되있는 것을 확인할 수 있다, **또 사진 위치가 이상한 것도 확인할 수 있다ㅜㅜ**, 이렇게 python으로 진짜 ppt를 다룰 수 있는게 신기하다, 신기한만큼 나에게 익숙하지 않은 느낌이고 효율이 안 좋은 느낌이다

### 6. 인치 대신 사용하는 단위(220714)
[python-pptx공식문서](https://python-pptx.readthedocs.io/en/latest/_modules/pptx/util.html#Centipoints)
를보니 `pptx.util` 에 `Inches,Cm,Emu` 등의 여러 단위가 있었다, 이제 과정들이 더 편해졌다

**슬라이드 크기도 바로 CM로 바꿧따^^**

### 7. 돌린 요소들 사이즈 맞춰주기(시도)
`4. 사진돌리기` 에서 나왔듯 세로가 더 긴 사진이면 90도를 돌리고 사이즈를 맞춰줬다, **이 때 사진을 돌리면 위치가 제대로 잡히지 않는다 ,아마 중간 축을 중심으로 사진이 돌아가고, 사이즈는 원본 사진대로 위치가 정해져서 사이즈가 안 맞는 것 같다**, 그래서 아래의 사진처럼 위치가 안 맞는걸 확인할 수 있다

![위치이상](https://user-images.githubusercontent.com/85085375/178867507-6e12109d-f54f-4c0d-ae12-c3c28c23d097.gif)

**그래서 `원래의 세로 - 바뀐 세로(90도 돌린 후)` 만큼의 gap을 맞춰줘야한다**

```py
if (height > width):
    pic.rotation = 90 # 돌리기
    width,height = height,width

    balance = width / height

    # 픽셀로 계산하면 안 맞기 때문에 height가 3.3cm 인 것을 활용
    heightCM = 3.3
    widthCM = round(balance * heightCM,1)

    gap = round(widthCM - heightCM,1) # 돌려서 난 차이
    print("gap",gap)

```

> 위 처럼 `gap`을 구했다, 이 `gap`을 그대로 `pic` 의 `left` 와 `top` 를 아래처럼 계산 했는데 자꾸 딱 안 맞고 넘어갔다

```py
pic.top = Cm((pic.top / 360000) - gap)
pic.left = Cm((pic.left / 360000) + gap)
```

![넘어감](https://user-images.githubusercontent.com/85085375/178877509-21986dc8-0fc3-4afb-aac8-6ab67389f697.png)

> 요기서 ㄹㅇ 삽질을 오지게 했다, **왜 안되는지 도저히 이해가 안 됐다**


### 8. 돌린 요소들 사이즈 맞춰주기(해결)
왜 안되는지 이해가 안 돼서 결국 그림을 그려봤는데, 바로 문제점을 찾았다, **이래서 이해가 안 되면 그려봐야 한다**

문제는 중간 축을 기준으로 돌렸기 때문에 반대에도 남는 부분이 생긴다, 즉,**남는 좌,우 부분을 합친게 `gap` 이고 위치를 맞추려면 `gap / 2` 를 해서 사용해야 한다**

![문제점](https://user-images.githubusercontent.com/85085375/178883640-e7a539f4-0604-40de-815e-2cbe8a47f920.png)

아래처럼 `gap` 을 반으로 나눴다

```py
pic.top = Cm((pic.top / 360000) - (gap/2))
pic.left = Cm((pic.left / 360000) + (gap/2))
```

![성공](https://user-images.githubusercontent.com/85085375/178884692-bb1613c4-d306-4046-812b-491783a53c3d.png)

> 위처럼 사진이 알맞은 위치에 위치하는걸 확인할 수 있다, 위에 살짝의 공백이 있는데, 계산하며 반올림을 해서 조금 오차가 생기는 것이다, **이렇게 수를 사용해서 실제로 만드는 일을 할 때 소수점을 적절히 다뤄야 하는게 짜증나는 것 같다**

### 9. 사진을 배치하기
`pointerLeft` 와 `pointerTop` 이라는 변수를 만들어서 이미지를 추가할 때 위치를 지정하는 원리로 배치를 한다

**사진 하나를 배치하고, 그 사진의 `width` 를 `pointerLeft` 에 더해준다, `pointerLeft` 가 다음 사진을 넣으면 슬라이드를 넘을 때 `0` 으로 초기화 해주고, `pointerTop` 을 `3.3cm` 추가해준다**

> 즉 `pointerLeft` 는 사진 추가를 끝내고 연산하고, `pointerTop` 은 사진을 넣고 바로 계산한다

#### 가로 사진 배치
위에 나온대로 사진을 배치 후 `pointerLeft`를 변경시켜준다

```py
pic = slide.shapes.add_picture(path+"\\"+name,Cm(pointerLeft),Cm(pointerTop),Cm(3.3)) # 3.3
.
.
.
pointerLeft = widthCM + pointerLeft
```

[사진배치](https://user-images.githubusercontent.com/85085375/178891011-29b7d02f-80f7-4342-89b9-a7ceeb5de705.png)

> 이 부분을 할 때 코드에서 논리적을 잘못된 부분(`heigthCM` 라는 변수가 `Cm` 함수로 감싸져서 Cm가 아닌)이 있어서 개발하는데 조금 힘들었다

#### 세로 사진 배치
사진을 넣고 바로 "사진을 넣으면 슬라이드를 넘어가는가?" 를 비교해야 했다

사진의 비율을 몰라서(사진을 돌릴지 말지 몰라서) 비효율 적이지만, 사진을 돌리는 로직에서 하면 위치 이동에 더 머리를 써야하는 일이 생기기 때문에, 그냥

1. 비율구하기
2. 비율에 따라 세로 또는 가로를 `pointerLeft` 에 추가
3. 슬라이드를 넘는지 판단

코드는 아래와 같다

```py
isOver = pointerLeft
if (height > width): # 세로가 더 클 땐 height를 추가
    isOver += 3.3 * (height / width)
else: # width를 추가
    isOver += 3.3

isOver = round(isOver,1)
print("isOver",isOver)

# 슬라이드를 넘으면
if (isOver > 29.7):
    pointerLeft = 0
    pointerTop += 3.3

    pic.left = Cm(pointerLeft)
    pic.top = Cm(pointerTop)
```

![결과](https://user-images.githubusercontent.com/85085375/178901218-c071d66c-22d0-4a49-8e03-771c336f88fc.png)

> 위 처럼 사진이 넘어서 추가되는 것을 확인할 수 있다, px을 다루는 과정이 조금 힘들었지만 비율을 활용해 계산하고 해결한 내가 자랑스럽다ㅋㅋ

### 10. 사진 변환
사실 변환할 사진을 아무거나 넣었을 때 아래처럼 예외로 처리되는 사진이 있다

![오류](https://user-images.githubusercontent.com/85085375/178902008-f8d4ff8d-9a6b-447c-b8b4-4954aab7d35f.png)

왜인진 정확하게 모르겠지만 [stackoverflow글을 보면](https://stackoverflow.com/questions/48712927/unable-to-insert-picture-in-slide)

```
이런 문제를 발생시키는 .jpg이미지가 실제로 스테레오 JPEG형식이기 때문에, 일반 JPEG형식으로 변환이 필요하다
```

라고 한다, 실제로 글 밑에 [링크되어 있었던 파일 형식 변환 사이트](https://image.online-convert.com/convert-to-jpg) 를 거치니 정상적으로 ppt에 들어갔었다

**Stackoverflow를 더 찾아보니 아래의 `PIL` 라이브러리의 `Image`를 사용해 해결이 가능하다 한다**

```py
from PIL import Image

for imagename in fileList:
    print(imagename)
    try:
        im = Image.open(path +"\\"+ imagename)
        im.save(path +"\\" + imagename)
    except:
        print(imagename, "안됨")
        pass

```

[Image.open](https://89douner.tistory.com/310)함수는 사진을 읽어서 정보를 어디 저장한다고 한다, 잘은 모르겠다^^

> 해치운건가.. 파이썬 자동화를 만들 때마다 드는 생각인데, 다 짠 코드를 보면 양이 생각보다 많지않다, 그만큼 파이썬이 간편하게 어떤 일을 하게 하는 것 같다, **자료가 적고, 영어고, 단위계산이 필요한 작업이라서 힘들었던 것 같다, 그래도 감을 잡으니 할만해졌다**