# 1. 필요한 메소드
- 엑셀 파일 관리
    - openExcel()
        - 엑셀 열기
    - createExcel()
        - 엑셀 없으면 생성
    - saveExcel()
        - 액셀 저장
- 시트 관리
    - getSheets()
        - 시트 목록 가져오기
    - openSheet(ws_name)
        - ws_name 시트 열기
    - createSheet(ws_name)
        - ws_name 시트 없으면 생성
    - initStatistics()
        - 엑셀 생성했을 때 Statistics 시트 초기화
    - initSheet()
        - 시트 생성했을 때 초기화
    > 시트 초기화시 [2. 셀 서식](#template) 참조
    - addMonthInStatistics()
        - 시트 생성했을 때 Statistics 시트에 해당 월 칼럼 추가

- 데이터 관리
    - getData(name, date)
        - 이름과 날짜에 해당하는 데이터 조회
    - inputData([dataDict](#dataDict))
        - 데이터 입력
    > #### 데이터 입력 시 주의할 부분
    > - 요일 판별하여 토, 일요일은 `근무 시간` 전체를 `특근 시간`으로 입력
    > - 평일일 경우 `공휴일(특근)` 체크되어 있을 시 공휴일로 판단, 주말과 같은 형태로 저장
    > - `잔업 시간`은 자동으로 계산하여 저장, `08:00` 이전, `17:00` 이후는 `잔업 시간`으로 입력
    > - `12:00 - 13:00`의 시각이 `근무 시각`에 포함되어 있는 `근무지`는 `점심 식사` 명목으로 `근무 시간`에서 `1` 제하고 입력
    > - `본사`에 `잔업 시간`이 있고, `저녁식사`가 체크되어 있을 시 `잔업 시간`에서 `0.5` 제하고 입력

- 기타
    - checkEmpList(ws_name)
        - 직원 목록 체크
    - addEmp(ws_name, name)
        - 시트에 직원 추가
    - getIndex(name, date)
        - 이름과 날짜에 해당하는 데이터 인덱스 조회

<span id='template'></span>
# 2. 셀 서식
<table style='text-align:center'>
    <thead>
        <tr>
            <th colspan=2></th>
            <th colspan=5 style='text-align:center; color:black; background-color:#F8CBAD;'>2020-11-01</th>
            <th colspan=5 style='text-align:center;'>2020-11-02</th>
        </tr>
        <tr>
            <th style='text-align:center'>이름</th>
            <th style='text-align:center'>근무지</th>
            <th style='text-align:center'>출퇴근 시각</th>
            <th style='text-align:center'>근무 시각</th>
            <th style='text-align:center'>근무 시간</th>
            <th style='text-align:center'>잔업 시간</th>
            <th style='text-align:center'>특근 시간</th>
            <th style='text-align:center'>출퇴근 시각</th>
            <th style='text-align:center'>근무 시각</th>
            <th style='text-align:center'>근무 시간</th>
            <th style='text-align:center'>잔업 시간</th>
            <th style='text-align:center'>특근 시간</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td rowspan=3>서충현</td>
            <td>본사</td>
            <td rowspan=3>07:00 - 21:30</td>
            <td>07:00 - 12:00</td>
            <td>5</td>
            <td></td>
            <td>5</td>
            <td rowspan=3>08:00 - 17:00</td>
            <td>08:00 - 14:30</td>
            <td>6.5</td>
            <td></td>
            <td></td>
        </tr>
        <tr>
            <td>400</td>
            <td>12:00 - 15:00</td>
            <td>3</td>
            <td></td>
            <td>3</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
        </tr>
        <tr>
            <td>어비리</td>
            <td>15:00 - 21:30</td>
            <td>6.5</td>
            <td></td>
            <td>6.5</td>
            <td>14:30 - 17:00</td>
            <td>2.5</td>
            <td></td>
            <td></td>
        </tr>
    </tbody>
</table>

<span id='dataDict'></span>
# 3. 데이터 딕셔너리
```
{
    'name': 이름,
    'date': {'y': 년, 'm': 월, 'd': 일},
    'totalWorking': 츨퇴근 시각(00:00 - 00:00),
    'hq': [본사 근무 시각, 본사 근무 시간, 본사 잔업 시간, 본사 특근 시간],
    '400': [400번지 근무 시각, 400번지 근무 시간, 400번지 잔업 시간, 400번지 특근 시간],
    'eobiri': [어비리 근무 시각, 어비리 근무 시간, 어비리 잔업 시간, 어비리 특근 시간]
}
```