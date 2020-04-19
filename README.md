# neis-school-code-mapper
시도별 교육청(17개) 에서 얻어온 학교명 목록 xlsx 를 neis 조회를 통해 neis school code(orgCode)를 mapping 시킵니다.

|파일명|교육청|
|--|--|
|busan-matcher.py|부산|
|chungbuk-matcher.py|충북|
|chungnam-matcher.py|충남|
|daegu-matcher.py|대구|
|daejeon-matcher.py|대전|
|gwangju-matcher.py|광주|
|gyeongbuk-matcher.py|경북|
|gyeongnam-matcher.py|경남|
|gyeonggi-matcher.py|경기|
|incheon-matcher.py|인천|
|jeju-matcher.py|제주|
|jeonbuk-matcher.py|전북|
|jeonnam-matcher.py|전남|
|kangwon-matcher.py|강원|
|sejong-matcher.py|세종|
|seoul-matcher.py|서울|
|ulsan-matcher.py|울산|

## 사용방법
1. 각 교육청 이름의 xlsx파일을 {교육청}-matcher.py와 동일한 경로에 위치한다.
2. 각 교육청 소속 고등학교 목록 파일 이름은 {교육청} 으로 맞춰야합니다.
3. 'python3 {교육청}-matcher.py' 실행


### xlxs format
|학교명(필수)|학제(옵션)|시도(옵션)|
|--|--|--|
|~~고등학교|고등학교|인천|
|~~고등학교|고등학교|인천|
|~~고등학교|고등학교|인천|
...

* 학교 이름이 A열에 있어야합니다. B, C열은 데이터가 없어도 됩니다.

---
### 결과
|--|학교명|학제|시도|
|--|--|--|--|
|C100000495|브니엘여자고등학교|고등학교|부산|
|C100000482|부산고등학교|고등학교|부산|
...




