# auto-make-picture-for-diary
다이어리에 붙일 사진을 ppt에 하나하나 넣는 과정이 귀찮아서 만들기 시작

## 1. `python-pptx` 설치
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

```