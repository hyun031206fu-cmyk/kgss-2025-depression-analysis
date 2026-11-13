# ==============================================================================
# 프로젝트: KGSS 2025 가족구조와 성인기 우울감 분석 재실행 프로젝트
# 연구 주제: 16세 당시 가족구조와 성인기 현재 우울감 사이의 관계
# 작성 목적: 신규 3개 가족구조 통합 유형 적용 및 전체 통계분석 재실행
# ==============================================================================

# ------------------------------------------------------------------------------
# STEP 0. 필수 라이브러리 로드
# ------------------------------------------------------------------------------
# 엑셀 파일을 읽고 쓰기 위한 라이브러리입니다.
# 만약 설치되어 있지 않다면 install.packages("readxl"), install.packages("writexl")을 먼저 실행하세요.
library(readxl)
library(writexl)

# 작업 폴더 내 '분석결과' 디렉터리가 없으면 자동 생성
if (!dir.exists("분석결과")) {
  dir.create("분석결과")
}

# ------------------------------------------------------------------------------
# STEP 1. 원본 데이터 불러오기 및 필요 변수 추출
# ------------------------------------------------------------------------------
# 원본 데이터 파일명: 2025_Data_Kor.xlsx
raw_file <- "2025_Data_Kor.xlsx"
data_raw <- read_excel(raw_file)

# 분석에 필요한 변수 목록:
# - FAMILY16: 16세 당시 동거 가족구조
# - WHYNOTPA23: 16세 당시 친부모와 동거하지 않은 주된 사유
# - CESD1 ~ CESD8: 성인기 우울감 측정 8개 문항
# - PAEDUC, MAEDUC: 부친 및 모친 최종 학력
# - AGE, SEX: 응답자 연령 및 성별
# - HOMPOP, RELATE2~6, AWAY2~6: 현재 가구 구성원 관계 및 동거 여부
# - FINALWT: 사후 가중치
sel_cols <- c("RESPID", "FAMILY16", "WHYNOTPA23",
              "CESD1", "CESD2", "CESD3", "CESD4",
              "CESD5", "CESD6", "CESD7", "CESD8",
              "PAEDUC", "MAEDUC",
              "AGE", "SEX",
              "HOMPOP",
              "RELATE2", "RELATE3", "RELATE4", "RELATE5", "RELATE6",
              "AWAY2", "AWAY3", "AWAY4", "AWAY5", "AWAY6",
              "FINALWT")

d <- as.data.frame(data_raw[, sel_cols])

# ------------------------------------------------------------------------------
# STEP 2. 결측치 및 특수 응답 코드 처리
# ------------------------------------------------------------------------------
# KGSS에서 -8(모름/무응답), -1(해당없음) 등 음수 코드는 통계 계산을 왜곡하므로 NA(결측) 처리합니다.
d[d < 0] <- NA

# ------------------------------------------------------------------------------
# STEP 3. 독립변수 생성 및 재코딩 (가장 핵심적인 변경 사항)
# ------------------------------------------------------------------------------
# (1) 이분형 독립변수 (nonintact):
#     0 = 친부모 가구
#     1 = 친부모와 함께 거주하지 않은 가구
d$nonintact <- ifelse(d$FAMILY16 == 1, 0, 1)

# (2) 기존 5개 가족구조 유형 (비교 및 검증용):
#     1: 친부모 모두와 거주
#     2: 친모와만 거주 (코드 5)
#     3: 친부와만 거주 (코드 4)
#     4: 재혼·계부모 가구 (코드 2, 3)
#     5: 친부모 모두와 비거주 (코드 6, 7, 8, 77)
d$famtype5 <- NA
d$famtype5[d$FAMILY16 == 1]                 <- "1.친부모"
d$famtype5[d$FAMILY16 == 5]                 <- "2.친어머니만"
d$famtype5[d$FAMILY16 == 4]                 <- "3.친아버지만"
d$famtype5[d$FAMILY16 %in% c(2, 3)]         <- "4.재혼(계부모)"
d$famtype5[d$FAMILY16 %in% c(6, 7, 8, 77)]  <- "5.친부모 모두 부재"
d$famtype5 <- factor(d$famtype5)

# (3) 신규 3개 가족구조 통합 유형 (요청 기준):
#     - 1.친부모 가구: 친부모 모두와 거주 (코드 1)
#     - 2.한부모 가구: 친모와만 거주(코드 5) + 친부와만 거주(코드 4)
#     - 3.부모 재구성·비동거 가구: 재혼·계부모(코드 2, 3) + 친부모 모두 비거주(코드 6, 7, 8, 77)
d$famtype3 <- NA
d$famtype3[d$FAMILY16 == 1]                 <- "1.친부모 가구"
d$famtype3[d$FAMILY16 %in% c(4, 5)]         <- "2.한부모 가구"
d$famtype3[d$FAMILY16 %in% c(2, 3, 6, 7, 8, 77)] <- "3.부모 재구성·비동거 가구"
d$famtype3 <- factor(d$famtype3, levels = c("1.친부모 가구", "2.한부모 가구", "3.부모 재구성·비동거 가구"))

# (4) 가족구조 형성 사유 (낙인 가설 검증용):
#     WHYNOTPA23: 1=사망, 2=이혼·별거, 3=직장/기타
d$reason <- NA
d$reason[d$FAMILY16 == 1]                            <- "0.친부모"
d$reason[d$FAMILY16 != 1 & d$WHYNOTPA23 %in% 1]      <- "1.사망"
d$reason[d$FAMILY16 != 1 & d$WHYNOTPA23 %in% 2]      <- "2.이혼·별거"
d$reason[d$FAMILY16 != 1 & !(d$WHYNOTPA23 %in% c(1, 2))] <- "3.기타·모름"
d$reason <- factor(d$reason)

# ------------------------------------------------------------------------------
# STEP 4. 종속변수 생성 (CES-D 성인기 우울감 점수)
# ------------------------------------------------------------------------------
# CES-D 8문항 중 긍정 문항(CESD4: 행복, CESD6: 즐거움)은 역코딩(5 - 점수)
d$CESD4r <- 5 - d$CESD4
d$CESD6r <- 5 - d$CESD6
d$cesd   <- d$CESD1 + d$CESD2 + d$CESD3 + d$CESD4r + d$CESD5 + d$CESD6r + d$CESD7 + d$CESD8

# 연구상 설정한 우울 고점수 집단(16점 이상, 임상 진단 아님)
d$dep_high <- ifelse(d$cesd >= 16, 1, 0)

# ------------------------------------------------------------------------------
# STEP 5. 통제변수 생성
# ------------------------------------------------------------------------------
# (1) 성별: 여성 1, 남성 0
d$female     <- ifelse(d$SEX == 2, 1, 0)
d$gender_kor <- ifelse(d$SEX == 2, "여성", "남성")

# (2) 연령 및 세대 구분
# 16세 도달 연도 ≈ 2025 - AGE + 16 = 2041 - AGE
# - 1.낙인세대(52세+): 1989년 이전 (조이혼율 1.1 이하, 강한 낙인)
# - 2.급증세대(37-51세): 1990~2004년 (외환위기 전후 이혼 급증)
# - 3.보편화세대(18-36세): 2005년 이후 (호주제 폐지 및 한부모 보편화)
d$age    <- d$AGE
d$year16 <- 2041 - d$AGE
d$gen <- cut(d$year16, breaks = c(-Inf, 1989, 2004, Inf),
             labels = c("1.낙인세대(52세+)", "2.급증세대(37-51세)", "3.보편화세대(18-36세)"))

# (3) 부모 교육수준: 학력 코드를 학교 다닌 연수로 변환 후 부모 중 최고 학력 적용
edu_years <- c(0, 6, 9, 12, 14, 16, 18, 21, 0)
d$pa_edu  <- edu_years[d$PAEDUC + 1]
d$ma_edu  <- edu_years[d$MAEDUC + 1]
d$par_edu <- pmax(d$pa_edu, d$ma_edu, na.rm = TRUE)

# (4) 현재 가구유형: 가구원 관계 및 실제 동거 여부 기준
rel <- d[, c("RELATE2", "RELATE3", "RELATE4", "RELATE5", "RELATE6")]
aw  <- d[, c("AWAY2", "AWAY3", "AWAY4", "AWAY5", "AWAY6")]
live_with <- function(codes) {
  rowSums(sapply(1:5, function(i) rel[[i]] %in% codes & aw[[i]] %in% 1)) > 0
}
spouse <- live_with(2)                 # 배우자와 동거
child  <- live_with(c(3, 4, 5))        # 자녀(손자녀 포함)와 동거
parent <- live_with(c(7, 8, 11))       # 부모·조부모와 동거

d$hhtype <- "5.기타(부모 동거 등)"
d$hhtype[spouse & child]  <- "1.부부+자녀"
d$hhtype[spouse & !child] <- "2.부부"
d$hhtype[!spouse & child] <- "4.한부모+자녀"
d$hhtype[parent]          <- "5.기타(부모 동거 등)"
d$hhtype[d$HOMPOP == 1]   <- "3.1인 가구"
d$hhtype <- factor(d$hhtype)

# ------------------------------------------------------------------------------
# STEP 6. 최종 분석 표본 확정 (Listwise Deletion: N = 1,156명)
# ------------------------------------------------------------------------------
vars_needed <- c("RESPID", "cesd", "dep_high", "nonintact", "famtype5", "famtype3", "reason",
                 "female", "gender_kor", "age", "gen", "par_edu", "hhtype", "FINALWT")

ana <- na.omit(d[, vars_needed])
cat(sprintf("전체 원자료 N = %d명 중 결측 55명을 제외한 최종 분석 표본 N = %d명 확정\n",
            nrow(d), nrow(ana)))

# 재분류 분석용 데이터 저장 (CSV 및 Excel)
write.csv(ana, "분석결과/분석용데이터_1156.csv", row.names = FALSE, fileEncoding = "UTF-8")
write_xlsx(ana, "분석결과/분석용데이터_1156.xlsx")

# ------------------------------------------------------------------------------
# STEP 7. 분석 실행: ① 기술통계 및 빈도분석
# ------------------------------------------------------------------------------
cat("\n=== [1] 기술통계 및 빈도분석 ===\n")
print(table(ana$famtype3))
print(prop.table(table(ana$famtype3)) * 100)
summary(ana$cesd)

# ------------------------------------------------------------------------------
# STEP 8. 분석 실행: ② 독립표본 t검정 (친부모 vs 비친부모)
# ------------------------------------------------------------------------------
cat("\n=== [2] 독립표본 t검정 ===\n")
tt <- t.test(cesd ~ nonintact, data = ana)
print(tt)

# ------------------------------------------------------------------------------
# STEP 9. 분석 실행: ③ 선형회귀분석 단계적 모형 (모형 1~4)
# ------------------------------------------------------------------------------
cat("\n=== [3] 선형회귀분석 단계적 모형 ===\n")
m1 <- lm(cesd ~ nonintact, data = ana)
m2 <- lm(cesd ~ nonintact + female + gen, data = ana)
m3 <- lm(cesd ~ nonintact + female + gen + hhtype, data = ana)
m4 <- lm(cesd ~ nonintact + female + gen + hhtype + par_edu, data = ana)

summary(m1)
summary(m2)
summary(m3)
summary(m4)

# ------------------------------------------------------------------------------
# STEP 10. 분석 실행: ④ 로지스틱 회귀분석 (고우울군 여부)
# ------------------------------------------------------------------------------
cat("\n=== [4] 로지스틱 회귀분석 (최종 모형) ===\n")
lg4 <- glm(dep_high ~ nonintact + female + gen + hhtype + par_edu, data = ana, family = binomial)
summary(lg4)
exp(coef(lg4))                       # 오즈비 (OR)
exp(confint.default(lg4))            # 95% 신뢰구간

# ------------------------------------------------------------------------------
# STEP 11. 분석 실행: ⑤ 세대별 상호작용 분석
# ------------------------------------------------------------------------------
cat("\n=== [5] 세대별 상호작용 분석 ===\n")
m_int <- lm(cesd ~ nonintact * gen + female + hhtype + par_edu, data = ana)
summary(m_int)
anova(m4, m_int)                     # 모형 간 증분 F검정

# ------------------------------------------------------------------------------
# STEP 12. 분석 실행: ⑥ 신규 3개 가족구조 유형별 회귀분석
# ------------------------------------------------------------------------------
cat("\n=== [6] 신규 3개 가족구조 유형별 회귀분석 (최종 모형) ===\n")
m_ft3_full <- lm(cesd ~ famtype3 + female + gen + hhtype + par_edu, data = ana)
summary(m_ft3_full)

# 한부모 vs 부모 재구성 집단 간 직접 차이 검정
ana$famtype3_rel <- relevel(ana$famtype3, ref = "2.한부모 가구")
m_ft3_diff <- lm(cesd ~ famtype3_rel + female + gen + hhtype + par_edu, data = ana)
summary(m_ft3_diff)

# ------------------------------------------------------------------------------
# STEP 13. 분석 실행: ⑦ 가족구조 형성 사유별 회귀분석
# ------------------------------------------------------------------------------
cat("\n=== [7] 가족구조 형성 사유별 회귀분석 ===\n")
m_rs_full <- lm(cesd ~ reason + female + gen + hhtype + par_edu, data = ana)
summary(m_rs_full)

# ------------------------------------------------------------------------------
# STEP 14. 분석 실행: ⑧ 가중치 적용 분석 및 ⑨ 연령 연속변수 강건성 분석
# ------------------------------------------------------------------------------
cat("\n=== [8] 가중치 적용 분석 (FINALWT) ===\n")
m4_wt    <- lm(cesd ~ nonintact + female + gen + hhtype + par_edu, data = ana, weights = FINALWT)
m_ft3_wt <- lm(cesd ~ famtype3 + female + gen + hhtype + par_edu, data = ana, weights = FINALWT)
summary(m4_wt)
summary(m_ft3_wt)

cat("\n=== [9] 연령 연속변수 강건성 분석 ===\n")
m4_age    <- lm(cesd ~ nonintact + female + age + hhtype + par_edu, data = ana)
m_ft3_age <- lm(cesd ~ famtype3 + female + age + hhtype + par_edu, data = ana)
summary(m4_age)
summary(m_ft3_age)

cat("\n모든 분석이 오류 없이 완료되었습니다! 결과표와 차트는 '분석결과' 폴더를 확인하세요.\n")
