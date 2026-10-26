# Power Automate - 정밀분석 미인계 예외 연동

MXKPCS는 새 SharePoint 열을 만들지 않고 `XRF_Requests`에 측정 건별 예외 감사 이벤트를 저장한다.

## DB 저장 규칙

- `Title`: `PREC_EXC_{Measurement_ID}`
- `field_1`: `PrecisionHandoffException`
- `field_2`: 적용 시 `Applied`, 해제 시 `Reverted`
- `field_3`: Item ID
- `field_4`: Trigger Measurement ID
- `field_5`: 예외 사유 코드
- `field_15`: 예외 상세 사유
- `Processed_At`, `Processed_By`: 최초/최근 예외 적용 정보
- `Reverted_At`, `Reverted_By`, `Revert_Reason`: 예외 해제 정보
- `Before_Target_Snapshot_JSON`: 상태, 사유, 근거자료와 감사정보 전체

예외는 품목 전체가 아니라 `XRF_Measurements.Title = XRF_Requests.field_4`인 한 측정 건에만 적용한다.

## Flow 변경

1. `Get_XRF_Requests`를 추가하고 `XRF_Requests` 목록을 읽는다.
2. `Select_PrecisionExceptions` 작업을 추가한다.

```text
eventType
item()?['field_1']

status
item()?['field_2']

triggerMeasurementId
item()?['field_4']
```

3. 정밀분석 메일 상태를 계산하기 직전에 현재 Measurement ID의 활성 예외를 찾는다.

```text
@and(
  equals(toLower(string(item()?['eventType'])), 'precisionhandoffexception'),
  equals(toUpper(string(item()?['status'])), 'APPLIED'),
  equals(string(item()?['triggerMeasurementId']), string(outputs('Compose_LatestMeasurementId')))
)
```

4. 활성 예외가 있으면 해당 Measurement ID를 `arrPrecisionMail`에 추가하지 않는다.

```text
@equals(length(body('Filter_ActivePrecisionException')), 0)
```

- `true`: 기존 정밀분석 상태 계산과 메일 배열 추가를 계속한다.
- `false`: 미인계 예외이므로 이 측정 건의 정밀분석 메일 추가를 건너뛴다.

## 일치 규칙

- XRF NG 원본 결과와 위험도는 변경하지 않는다.
- 정기 주기와 이행 상태는 변경하지 않는다.
- MXKPCS의 정밀분석 탭과 위험도 분석 Flow에서는 `미인계 예외` 이력으로 집계한다.
- `정밀분석 진행중` 건수에서는 제외한다.
- 예외 해제 시 `field_2=Reverted`가 되므로 다음 Flow 실행부터 다시 메일 대상이 된다.
- 이미 정밀분석에 인계되었거나 결과가 등록된 Measurement는 미인계 예외처리를 차단한다.
