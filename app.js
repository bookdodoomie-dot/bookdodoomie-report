// app.js

document.addEventListener('DOMContentLoaded', () => {
    // 1. 오늘 날짜 자동 표기 설정
    initDate();

    // 2. 폼 바인딩 및 실시간 리포트 업데이트 활성화
    initBindings();

    // 3. 우측 A4 미리보기 컨테이너 동적 축소/확대(Scale) 최적화
    initScaling();

    // 4. 엑셀 데이터 빠른 입력 및 자동 소견 매칭 초기화
    initExcelImport();
});

/**
 * 1. 오늘 날짜를 한국형 및 테이블형으로 초기화
 */
function initDate() {
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');

    const headerDate = `${year}-${month}-${day}`;
    const footerDate = `${year}년 ${month}월 ${day}일`;

    const viewDateEl = document.getElementById('view-date');
    const viewFooterDateEl = document.getElementById('view-footer-date');

    if (viewDateEl) viewDateEl.textContent = headerDate;
    if (viewFooterDateEl) viewFooterDateEl.textContent = footerDate;
}

/**
 * 2. 실시간 데이터 양방향 바인딩 및 이벤트 리스너 연동
 */
function initBindings() {
    // A. 학생 기본 정보 연동
    bindTextInput('student-name', 'view-name');
    bindTextInput('academy-name', 'view-center');
    bindTextInput('student-grade', 'view-grade');
    
    // 독서 단계와 세부 구분 결합 실시간 바인딩
    const levelSelect = document.getElementById('student-reading-level');
    const sublevelSelect = document.getElementById('student-reading-sublevel');
    const viewReadingLevel = document.getElementById('view-reading-level');

    const updateReportReadingLevel = () => {
        if (levelSelect && sublevelSelect && viewReadingLevel) {
            const levelVal = levelSelect.value;
            const sublevelVal = sublevelSelect.value;
            if (sublevelVal === "기본") {
                viewReadingLevel.textContent = levelVal;
            } else {
                viewReadingLevel.textContent = `${levelVal}(${sublevelVal})`;
            }
        }
    };

    if (levelSelect && sublevelSelect) {
        levelSelect.addEventListener('change', updateReportReadingLevel);
        sublevelSelect.addEventListener('change', updateReportReadingLevel);
        // 초기화 데이터 동기화
        updateReportReadingLevel();
    }

    // B. 영역별 세부 의견 기술 연동
    bindTextInput('comment-vocab', 'view-vocab-comment');
    bindTextInput('comment-analysis', 'view-analysis-comment');
    bindTextInput('comment-inference', 'view-inference-comment');
    bindTextInput('comment-critical', 'view-critical-comment');
    bindTextInput('comment-meta', 'view-meta-comment');
    bindTextInput('comment-overall', 'view-overall-comment');

    // C. 점수 입력 및 차트 바인딩
    const domains = [
        { key: 'vocab', name: '어휘력' },
        { key: 'analysis', name: '분석력' },
        { key: 'inference', name: '추론력' },
        { key: 'critical', name: '비판력' },
        { key: 'meta', name: '메타인지' }
    ];

    domains.forEach(domain => {
        const slider = document.getElementById(`score-${domain.key}`);
        const number = document.getElementById(`num-${domain.key}`);

        if (slider && number) {
            // 슬라이더 값 변경 시 이벤트
            slider.addEventListener('input', (e) => {
                const val = parseInt(e.target.value, 10) || 0;
                number.value = val;
                updateScoreAndChart(domain.key, val);
            });

            // 숫자 직접 입력 변경 시 이벤트
            number.addEventListener('input', (e) => {
                let val = parseInt(e.target.value, 10);
                if (isNaN(val)) val = 0;
                if (val < 0) val = 0;
                if (val > 100) val = 100;
                
                slider.value = val;
                updateScoreAndChart(domain.key, val);
            });

            // 초기 렌더링 값 적용
            updateScoreAndChart(domain.key, parseInt(slider.value, 10));
        }
    });

    // D. 인쇄 및 PDF 저장 버튼 이벤트 연동
    const btnPrint = document.getElementById('btn-print');
    const btnPdf = document.getElementById('btn-pdf');

    if (btnPrint) {
        btnPrint.addEventListener('click', () => {
            window.print();
        });
    }

    if (btnPdf) {
        btnPdf.addEventListener('click', () => {
            // 브라우저의 PDF 내보내기 창을 직접 호출합니다.
            window.print();
        });
    }
}

/**
 * 텍스트 입력 소스 엘리먼트와 리포트 표기 엘리먼트 실시간 매핑 함수
 */
function bindTextInput(inputId, viewId) {
    const inputEl = document.getElementById(inputId);
    const viewEl = document.getElementById(viewId);

    if (inputEl && viewEl) {
        inputEl.addEventListener('input', (e) => {
            // 줄바꿈이 있는 텍스트는 <br> 혹은 단락으로 보존되도록 처리
            if (inputEl.tagName.toLowerCase() === 'textarea') {
                viewEl.innerHTML = e.target.value.replace(/\n/g, '<br>');
            } else {
                viewEl.textContent = e.target.value;
            }
        });
        
        // 초기화 데이터 동기화
        if (inputEl.tagName.toLowerCase() === 'textarea') {
            viewEl.innerHTML = inputEl.value.replace(/\n/g, '<br>');
        } else {
            viewEl.textContent = inputEl.value;
        }
    }
}

/**
 * 점수대에 따른 등급 분류 기준
 * 80점 이상 ➡️ "상위권"
 * 70점 ~ 79점 ➡️ "평균 이상"
 * 60점 ~ 69점 ➡️ "평균"
 * 59점 이하 ➡️ "보완 필요"
 */
function getRating(score) {
    if (score >= 80) {
        return { label: '상위권', className: 'rating-high' };
    } else if (score >= 70) {
        return { label: '평균 이상', className: 'rating-good' };
    } else if (score >= 60) {
        return { label: '평균', className: 'rating-normal' };
    } else {
        return { label: '보완 필요', className: 'rating-low' };
    }
}

/**
 * 특정 영역의 점수가 변경될 때 수치 표기, 배지 색상 및 SVG 차트를 갱신
 */
function updateScoreAndChart(domainKey, score) {
    // 1. 테이블 수치 변경
    const valEl = document.getElementById(`val-${domainKey}`);
    if (valEl) valEl.textContent = score;

    // 2. 점수대별 평가 판단
    const rating = getRating(score);

    // 3. 좌측 입력폼 배지 갱신
    const inputBadge = document.getElementById(`badge-${domainKey}`);
    if (inputBadge) {
        inputBadge.textContent = rating.label;
        // 클래스 초기화 후 알맞은 등급 적용
        inputBadge.className = 'badge';
        inputBadge.classList.add(rating.className);
    }

    // 4. 우측 리포트 테이블 배지 갱신
    const reportBadge = document.getElementById(`rate-${domainKey}`);
    if (reportBadge) {
        reportBadge.textContent = rating.label;
        reportBadge.className = 'badge';
        reportBadge.classList.add(rating.className);
    }

    // 5. 실시간 SVG 방사형 차트 갱신 연산
    updateRadarChart();

    // 6. 점수 변경에 따른 좌측 입력 소견 자동 갱신 (선생님이 직접 편집하여 커스텀한 글이 아닐 때만 동적 반영)
    const textarea = document.getElementById(`comment-${domainKey}`);
    if (textarea) {
        const currentVal = textarea.value.trim();
        // 비어있거나 혹은 이미 다른 자동 완성 텍스트 중 하나인 경우
        const isAuto = currentVal === "" || AUTO_COMMENTS[domainKey].some(item => item.text.trim() === currentVal);
        if (isAuto) {
            const nextComment = getAutoCommentText(domainKey, score);
            textarea.value = nextComment;
            // 리포트 뷰 갱신 발동
            const viewEl = document.getElementById(`view-${domainKey}-comment`);
            if (viewEl) {
                viewEl.innerHTML = nextComment.replace(/\n/g, '<br>');
            }
        }
    }

    // 7. 종합 처방 자동 갱신 (마찬가지로 자동완성 형태일 때만 동적 반영)
    const overallTextarea = document.getElementById('comment-overall');
    if (overallTextarea) {
        const studentName = document.getElementById('student-name').value.trim() || "학생";
        const currentOverall = overallTextarea.value.trim();
        
        // 현재 화면상의 5대 점수 취득
        const currentScores = {
            vocab: parseInt(document.getElementById('score-vocab').value, 10) || 0,
            analysis: parseInt(document.getElementById('score-analysis').value, 10) || 0,
            inference: parseInt(document.getElementById('score-inference').value, 10) || 0,
            critical: parseInt(document.getElementById('score-critical').value, 10) || 0,
            meta: parseInt(document.getElementById('score-meta').value, 10) || 0
        };
        
        const isAutoOverall = currentOverall === "" || currentOverall.includes("학생은 영역 평가 항목 중");
        if (isAutoOverall) {
            const nextOverall = getAutoOverallComment(studentName, currentScores);
            overallTextarea.value = nextOverall;
            const overallView = document.getElementById('view-overall-comment');
            if (overallView) {
                overallView.innerHTML = nextOverall.replace(/\n/g, '<br>');
            }
        }
    }
}

/**
 * 실시간 SVG 거미줄(방사형) 차트 수학적 렌더링
 * 중심점 (160, 160)
 * 최대 축 반경 R = 100
 * 다섯 개 축의 벡터 방향 매핑:
 *  - 어휘력 (Vocab)       : -90도 (위)
 *  - 분석력 (Analysis)    : -18도 (우상)
 *  - 추론력 (Inference)   : 54도 (우하)
 *  - 비판력 (Critical)    : 126도 (좌하)
 *  - 메타인지 (Meta)      : 198도 (좌상)
 */
function updateRadarChart() {
    const cx = 160;
    const cy = 160;
    const maxR = 100;

    // 현재 각 폼 필드의 점수 취득
    const scoreVocab = parseInt(document.getElementById('score-vocab').value, 10) || 0;
    const scoreAnalysis = parseInt(document.getElementById('score-analysis').value, 10) || 0;
    const scoreInference = parseInt(document.getElementById('score-inference').value, 10) || 0;
    const scoreCritical = parseInt(document.getElementById('score-critical').value, 10) || 0;
    // 메타인지 점수는 슬라이더에서 가져옵니다.
    const metaSlider = document.getElementById('score-meta');
    const scoreMeta = metaSlider ? (parseInt(metaSlider.value, 10) || 0) : 60;

    // 각 영역별 반경 환산
    const rVocab = (scoreVocab / 100) * maxR;
    const rAnalysis = (scoreAnalysis / 100) * maxR;
    const rInference = (scoreInference / 100) * maxR;
    const rCritical = (scoreCritical / 100) * maxR;
    const rMeta = (scoreMeta / 100) * maxR;

    // 5대 축 각도 정의 (라디안)
    const angleVocab = -Math.PI / 2;             // -90도 (위)
    const angleAnalysis = -18 * Math.PI / 180;    // -18도 (우상)
    const angleInference = 54 * Math.PI / 180;    // 54도 (우하)
    const angleCritical = 126 * Math.PI / 180;   // 126도 (좌하)
    const angleMeta = 198 * Math.PI / 180;       // 198도 (좌상)

    // 각 점의 X, Y 좌표 수학적 산출
    const xVocab = cx + rVocab * Math.cos(angleVocab);
    const yVocab = cy + rVocab * Math.sin(angleVocab);

    const xAnalysis = cx + rAnalysis * Math.cos(angleAnalysis);
    const yAnalysis = cy + rAnalysis * Math.sin(angleAnalysis);

    const xInference = cx + rInference * Math.cos(angleInference);
    const yInference = cy + rInference * Math.sin(angleInference);

    const xCritical = cx + rCritical * Math.cos(angleCritical);
    const yCritical = cy + rCritical * Math.sin(angleCritical);

    const xMeta = cx + rMeta * Math.cos(angleMeta);
    const yMeta = cy + rMeta * Math.sin(angleMeta);

    // A. 다각형 그리기 (points 어트리뷰트 업데이트)
    const polyPoints = `${xVocab},${yVocab} ${xAnalysis},${yAnalysis} ${xInference},${yInference} ${xCritical},${yCritical} ${xMeta},${yMeta}`;
    const dataPoly = document.getElementById('data-poly');
    if (dataPoly) {
        dataPoly.setAttribute('points', polyPoints);
    }

    // B. 다각형 꼭짓점 마커 동적 갱신
    setCircleAttr('dot-vocab', xVocab, yVocab);
    setCircleAttr('dot-analysis', xAnalysis, yAnalysis);
    setCircleAttr('dot-inference', xInference, yInference);
    setCircleAttr('dot-critical', xCritical, yCritical);
    setCircleAttr('dot-meta', xMeta, yMeta);
}

function setCircleAttr(id, cx, cy) {
    const el = document.getElementById(id);
    if (el) {
        el.setAttribute('cx', cx);
        el.setAttribute('cy', cy);
    }
}

/**
 * 3. 우측 A4 미리보기 화면 자동 크기 맞춤 (Scale) 함수 연동
 * 브라우저 크기를 조절하더라도 210mm x 297mm의 거대한 A4용지가 깨지거나 잘리지 않고,
 * 우측 결과 영역 박스 비율에 완벽히 피트되어 한눈에 보이게 돕습니다.
 */
function initScaling() {
    // 초기 로딩 후 스케일 조절 수행
    setTimeout(fitA4Preview, 50);

    // 윈도우 크기 변경 시 및 디바이스 방향 변경 시 실시간 최적화
    window.addEventListener('resize', fitA4Preview);
    
    // 혹시 모를 폰트 로딩 등으로 인한 레이아웃 오차 예방을 위해 로드 후 한번 더 실행
    window.addEventListener('load', fitA4Preview);
}

function fitA4Preview() {
    const previewPane = document.querySelector('.preview-pane');
    const container = document.querySelector('.a4-container');
    const page = document.getElementById('printable-report');
    
    if (!previewPane || !container || !page) return;
    
    // A4 실제 설계 타겟 픽셀 (96dpi 기준 794px x 1123px)
    const targetW = 794;
    const targetH = 1123;
    
    // 미리보기 패널의 패딩 여유 여백 반영 크기 (좌우상하 최소 40px 여백 확보)
    const padding = 40;
    const availableW = previewPane.clientWidth - padding;
    const availableH = previewPane.clientHeight - padding;
    
    // 스케일 가로세로 비율 중 최솟값 선정 (절대 1을 넘지 않도록 하여 본래 고해상도 인쇄 품질 보존)
    const scaleW = availableW / targetW;
    const scaleH = availableH / targetH;
    let scale = Math.min(scaleW, scaleH);
    
    // 0.1보다 작아지지 않도록 안전장치 설정
    if (scale < 0.1) scale = 0.1;
    if (scale > 1.0) scale = 1.0;
    
    // CSS 3D 가속 변환 적용 (트랜스폼 중심점: 상단 중앙)
    page.style.transform = `scale(${scale})`;
    page.style.transformOrigin = 'top center';
    
    // 부모 컨테이너의 가상 감싸기 너비/높이를 스케일링된 실제 높이와 매칭시켜 스크롤바가 정확히 동작하도록 제어
    container.style.height = `${targetH * scale}px`;
    container.style.width = `${targetW * scale}px`;
}

// ==========================================
// 4. 엑셀 데이터 파싱 및 실시간 자동 리포트 로더
// ==========================================

let importedStudents = []; // 엑셀에서 파싱된 학생 리스트 전역 보관

/**
 * 엑셀 가져오기 및 파싱 기능 초기화
 */
function initExcelImport() {
    const dropzone = document.getElementById('excel-dropzone');
    const fileInput = document.getElementById('excel-file-input');
    const studentSelect = document.getElementById('excel-student-select');

    if (!dropzone || !fileInput) return;

    // 드롭존 클릭 이벤트
    dropzone.addEventListener('click', () => {
        fileInput.click();
    });

    // 드래그 앤 드롭 애니메이션 효과 바인딩
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleExcelFile(files[0]);
        }
    });

    // 파일 선택창 변경 이벤트
    fileInput.addEventListener('change', (e) => {
        const files = e.target.files;
        if (files.length > 0) {
            handleExcelFile(files[0]);
        }
    });

    // 학생 선택 변경 시 양방향 폼 바인딩 동기화
    if (studentSelect) {
        studentSelect.addEventListener('change', (e) => {
            const index = parseInt(e.target.value, 10);
            if (!isNaN(index) && importedStudents[index]) {
                loadStudentData(importedStudents[index]);
            }
        });
    }
}

/**
 * 엑셀 파일 바이너리 로드 및 SheetJS 파싱 (하이브리드 모드 지원)
 */
function handleExcelFile(file) {
    const fileNameDisplay = document.getElementById('file-name-display');
    if (fileNameDisplay) {
        fileNameDisplay.textContent = file.name;
        fileNameDisplay.classList.add('loaded');
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            importedStudents = [];
            
            // 시트가 여러 개이고, D18 셀 주소(특정 서식 구조)가 활성화되어 있는지 자동 판별
            let isFormStyle = false;
            
            // 엑셀에서 정수값 파싱하는 로컬 헬퍼
            const parseScoreValue = (val) => {
                if (val === undefined || val === null) return NaN;
                const str = String(val).trim();
                if (str === "") return NaN;
                const digits = str.replace(/[^0-9]/g, '');
                if (digits === "") return NaN;
                return parseInt(digits, 10);
            };

            // 모든 시트를 스캔하여 D18, F18, H18, J18 중 하나에라도 숫자가 입력되어 있다면 폼 서식(Worksheet당 1명 학생) 스타일로 간주
            for (let i = 0; i < workbook.SheetNames.length; i++) {
                const sheetName = workbook.SheetNames[i];
                const worksheet = workbook.Sheets[sheetName];
                if (worksheet) {
                    const d18 = worksheet['D18'] ? worksheet['D18'].v : null;
                    const f18 = worksheet['F18'] ? worksheet['F18'].v : null;
                    const h18 = worksheet['H18'] ? worksheet['H18'].v : null;
                    const j18 = worksheet['J18'] ? worksheet['J18'].v : null;
                    
                    const d18Num = parseScoreValue(d18);
                    const f18Num = parseScoreValue(f18);
                    const h18Num = parseScoreValue(h18);
                    const j18Num = parseScoreValue(j18);
                    
                    if (!isNaN(d18Num) || !isNaN(f18Num) || !isNaN(h18Num) || !isNaN(j18Num)) {
                        isFormStyle = true;
                        break;
                    }
                }
            }
            
            // 1. 폼 서식 스타일 파싱 모드 (시트 = 학생 한 명, 점수는 D18, F18, H18, J18)
            if (isFormStyle) {
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    if (!worksheet) return;
                    
                    // 시트 이름 자체가 학생 이름이거나, 혹은 시트 내부에서 학생 이름을 찾음
                    let studentName = sheetName.trim();
                    
                    // 어휘력 D18, 분석력 F18, 추론력 H18, 비판력 J18 직접 조회 및 파싱
                    const scoreVocab = parseScore(worksheet['D18']);
                    const scoreAnalysis = parseScore(worksheet['F18']);
                    const scoreInference = parseScore(worksheet['H18']);
                    const scoreCritical = parseScore(worksheet['J18']);
                    
                    // 학원명/학년 또한 시트의 특정 고정 셀에서 찾아보기 (기본값 설정 후 광범위 셀 스캔)
                    let academy = "꿈동산지역아동센터";
                    let grade = "2026년 초4";
                    
                    // 일반적인 서식 셀 스캔 유틸 ( raw 값 .v 가 없으면 포맷팅값 .w 라도 조회 )
                    const getCellValue = (cellAddr) => {
                        const cell = worksheet[cellAddr];
                        if (!cell) return "";
                        if (cell.v !== undefined && cell.v !== null) return String(cell.v).trim();
                        if (cell.w !== undefined && cell.w !== null) return String(cell.w).trim();
                        return "";
                    };
                    
                    // 서식 스캐너로 상단 영역(행 3~6, 열 B~G)에서 학원명, 학년, 학생이름 정밀 탐색
                    const scanCells = [
                        'C3', 'D3', 'E3', 'F3', 'G3',
                        'B4', 'C4', 'D4', 'E4', 'F4', 'G4',
                        'B5', 'C5', 'D5', 'E5', 'F5', 'G5',
                        'B6', 'C6', 'D6', 'E6', 'F6', 'G6'
                    ];
                    
                    let nameCandidate = "";
                    for (const cell of scanCells) {
                        const val = getCellValue(cell);
                        if (!val) continue;
                        
                        // 기관명 추출
                        if (val.includes("학원") || val.includes("센터") || val.includes("아동") || val.includes("교실") || val.includes("공부방")) {
                            academy = val;
                        }
                        
                        // 학년 추출 (단, 일자는 패스)
                        if (val.includes("학년") || val.includes("초") || val.includes("중") || val.includes("고") || val.includes("202")) {
                            if (!val.includes("-") && !val.includes("/")) {
                                grade = val;
                            }
                        }
                        
                        // 이름 추출 후보 (순수 한글 2~4글자)
                        const nameRegex = /^[가-힣]{2,4}$/;
                        if (nameRegex.test(val)) {
                            const excludeWords = ["학년", "학원", "센터", "이름", "성명", "구분", "날짜", "일자", "점수", "평가", "총평", "소견", "어휘", "분석", "추론", "비판"];
                            if (!excludeWords.some(w => val.includes(w))) {
                                nameCandidate = val;
                            }
                        }
                    }
                    
                    if (nameCandidate) {
                        studentName = nameCandidate;
                    }
                    
                    const student = {
                        name: studentName,
                        academy: academy,
                        grade: grade,
                        scores: {
                            vocab: scoreVocab,
                            analysis: scoreAnalysis,
                            inference: scoreInference,
                            critical: scoreCritical,
                            meta: 60
                        },
                        comments: {
                            vocab: "",
                            analysis: "",
                            inference: "",
                            critical: "",
                            meta: "",
                            overall: ""
                        }
                    };
                    
                    // 소견 자동 생성
                    generateAutoComments(student);
                    importedStudents.push(student);
                });
            } 
            // 2. 일반 테이블 형식 파싱 모드 (기존 구현)
            else {
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                if (rows.length < 2) {
                    alert("엑셀 파일에 분석 가능한 데이터 행이 없습니다.");
                    return;
                }

                const headers = rows[0];
                const colIdx = {
                    name: findColumnIndex(headers, ['이름', '학생', '성명', 'name', 'student']),
                    academy: findColumnIndex(headers, ['학원', '아동센터', '기관', '소속', '센터', 'academy', 'center']),
                    grade: findColumnIndex(headers, ['학년', '학급', '반', 'grade', 'class', '구분']),
                    vocab: findColumnIndex(headers, ['어휘', 'vocab']),
                    analysis: findColumnIndex(headers, ['분석', 'analysis']),
                    inference: findColumnIndex(headers, ['추론', 'infer']),
                    critical: findColumnIndex(headers, ['비판', 'critic']),
                    meta: findColumnIndex(headers, ['메타', 'meta']),
                    commentVocab: findColumnIndex(headers, ['어휘력세부', '어휘세부', '어휘의견', '어휘평가', '어휘코멘트']),
                    commentAnalysis: findColumnIndex(headers, ['분석력세부', '분석세부', '분석의견', '분석평가', '분석코멘트']),
                    commentInference: findColumnIndex(headers, ['추론력세부', '추론세부', '추론의견', '추론평가', '추론코멘트']),
                    commentCritical: findColumnIndex(headers, ['비판력세부', '비판세부', '비판의견', '비판평가', '비판코멘트']),
                    commentMeta: findColumnIndex(headers, ['메타인지세부', '메타세부', '메타의견', '메타평가', '메타코멘트', '메타인지코멘트']),
                    commentOverall: findColumnIndex(headers, ['종합', '총평', '소견', '처방', '제언', 'overall'])
                };

                if (colIdx.name === -1) {
                    alert("엑셀 파일 내에서 '학생 이름' 헤더 컬럼을 매핑하지 못했습니다.");
                    return;
                }

                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row || row.length === 0) continue;

                    const nameVal = row[colIdx.name];
                    if (!nameVal || String(nameVal).trim() === '') continue;

                    const student = {
                        name: String(nameVal).trim(),
                        academy: colIdx.academy !== -1 && row[colIdx.academy] ? String(row[colIdx.academy]).trim() : "꿈동산지역아동센터",
                        grade: colIdx.grade !== -1 && row[colIdx.grade] ? String(row[colIdx.grade]).trim() : "2026년 초4",
                        scores: {
                            vocab: parseScore(colIdx.vocab !== -1 ? row[colIdx.vocab] : 70),
                            analysis: parseScore(colIdx.analysis !== -1 ? row[colIdx.analysis] : 70),
                            inference: parseScore(colIdx.inference !== -1 ? row[colIdx.inference] : 70),
                            critical: parseScore(colIdx.critical !== -1 ? row[colIdx.critical] : 70),
                            meta: parseScore(colIdx.meta !== -1 ? row[colIdx.meta] : 60)
                        },
                        comments: {
                            vocab: colIdx.commentVocab !== -1 && row[colIdx.commentVocab] ? String(row[colIdx.commentVocab]).trim() : "",
                            analysis: colIdx.commentAnalysis !== -1 && row[colIdx.commentAnalysis] ? String(row[colIdx.commentAnalysis]).trim() : "",
                            inference: colIdx.commentInference !== -1 && row[colIdx.commentInference] ? String(row[colIdx.commentInference]).trim() : "",
                            critical: colIdx.commentCritical !== -1 && row[colIdx.commentCritical] ? String(row[colIdx.commentCritical]).trim() : "",
                            meta: colIdx.commentMeta !== -1 && row[colIdx.commentMeta] ? String(row[colIdx.commentMeta]).trim() : "",
                            overall: colIdx.commentOverall !== -1 && row[colIdx.commentOverall] ? String(row[colIdx.commentOverall]).trim() : ""
                        }
                    };

                    generateAutoComments(student);
                    importedStudents.push(student);
                }
            }

            if (importedStudents.length === 0) {
                alert("불러올 수 있는 유효한 학생 데이터가 없습니다.");
                return;
            }

            // 드롭다운 갱신 및 표시
            const studentSelect = document.getElementById('excel-student-select');
            const selectGroup = document.getElementById('student-select-group');
            
            if (studentSelect && selectGroup) {
                studentSelect.innerHTML = '';
                importedStudents.forEach((st, idx) => {
                    const option = document.createElement('option');
                    option.value = idx;
                    option.textContent = `${st.name} (${st.academy} - ${st.grade})`;
                    studentSelect.appendChild(option);
                });
                selectGroup.style.display = 'block';
            }

            // 첫 번째 학생 자동 로딩
            loadStudentData(importedStudents[0]);

        } catch (error) {
            console.error(error);
            alert("엑셀 파일을 파싱하는 데 실패했습니다. 파일 형식을 다시 점검해 주세요.");
        }
    };
    reader.readAsArrayBuffer(file);
}

/**
 * 정규화 기반의 Fuzzy Header Column 매칭 유틸리티
 */
function findColumnIndex(headers, keywords) {
    for (let i = 0; i < headers.length; i++) {
        if (headers[i] === undefined || headers[i] === null) continue;
        // 공백과 대소문자를 제거하여 매치 강도 극대화
        const headerStr = String(headers[i]).replace(/\s+/g, '').toLowerCase();
        for (const kw of keywords) {
            const normalizedKw = kw.replace(/\s+/g, '').toLowerCase();
            if (headerStr.includes(normalizedKw)) {
                return i;
            }
        }
    }
    return -1;
}

/**
 * 안전한 정수 변환 및 0-100 캡핑
 */
function parseScore(val) {
    if (val === undefined || val === null) return 70;
    
    let rawVal = val;
    // 만약 SheetJS 셀 객체라면 내부 값 꺼내기
    if (typeof val === 'object' && val !== null) {
        if (val.v !== undefined && val.v !== null) rawVal = val.v;
        else if (val.w !== undefined && val.w !== null) rawVal = val.w;
        else return 70;
    }
    
    const str = String(rawVal).trim();
    const digits = str.replace(/[^0-9]/g, '');
    let num = parseInt(digits, 10);
    if (isNaN(num)) {
        num = parseInt(str, 10);
    }
    if (isNaN(num)) num = 70;
    return Math.min(Math.max(num, 0), 100);
}

function loadStudentData(student) {
    // 1. 일반 텍스트 정보 바인딩
    setInputValue('student-name', student.name);
    setInputValue('academy-name', student.academy);
    setInputValue('student-grade', student.grade);
    
    // 독서 단계 기본값(1단계) 혹은 학생 데이터에 들어있는 값 주입 (엑셀 제외 수동 기본 세팅)
    setInputValue('student-reading-level', student.readingLevel || "1단계");
    setInputValue('student-reading-sublevel', student.readingSublevel || "기본");

    // 2. 5대 영역 점수 수치 연동
    setInputValue('score-vocab', student.scores.vocab);
    setInputValue('num-vocab', student.scores.vocab);
    
    setInputValue('score-analysis', student.scores.analysis);
    setInputValue('num-analysis', student.scores.analysis);
    
    setInputValue('score-inference', student.scores.inference);
    setInputValue('num-inference', student.scores.inference);
    
    setInputValue('score-critical', student.scores.critical);
    setInputValue('num-critical', student.scores.critical);
    
    // 메타인지 점수 슬라이더 값 바인딩
    const metaScore = student.scores.meta !== undefined ? student.scores.meta : 60;
    setInputValue('score-meta', metaScore);
    setInputValue('num-meta', metaScore);

    // 3. 영역별 소견 및 종합 평가 텍스트에 할당
    setInputValue('comment-vocab', student.comments.vocab);
    setInputValue('comment-analysis', student.comments.analysis);
    setInputValue('comment-inference', student.comments.inference);
    setInputValue('comment-critical', student.comments.critical);
    setInputValue('comment-meta', student.comments.meta || "");
    setInputValue('comment-overall', student.comments.overall);

    // 4. 모든 입력폼에서 'input' 또는 'change' 이벤트를 발동시켜 우측 리포트 바인딩 및 렌더링 동기 유도
    const inputIds = [
        'student-name', 'academy-name', 'student-grade', 
        'student-reading-level', 'student-reading-sublevel',
        'score-vocab', 'score-analysis', 'score-inference', 'score-critical', 'score-meta',
        'comment-vocab', 'comment-analysis', 'comment-inference', 'comment-critical', 'comment-meta',
        'comment-overall'
    ];

    inputIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.dispatchEvent(new Event('input'));
            el.dispatchEvent(new Event('change'));
        }
    });

    // 5. 점수 뱃지 상태 및 SVG 차트 업데이트 재동기화
    const domains = ['vocab', 'analysis', 'inference', 'critical', 'meta'];
    domains.forEach(d => {
        const scoreVal = d === 'meta' ? metaScore : student.scores[d];
        updateScoreAndChart(d, scoreVal);
    });
}

function setInputValue(id, val) {
    const el = document.getElementById(id);
    if (el) {
        el.value = val;
    }
}

/**
 * 엑셀 데이터에 공란으로 구성된 소견 및 제언을 전문가 등급의 피드백으로 자동 설계 생성
 */
function generateAutoComments(student) {
    const s = student.scores;
    const c = student.comments;

    // B. 비어있는 개별 영역 코멘트 자동 대체 주입 (6등급 프리미엄 엔진)
    if (!c.vocab) c.vocab = getAutoCommentText('vocab', s.vocab);
    if (!c.analysis) c.analysis = getAutoCommentText('analysis', s.analysis);
    if (!c.inference) c.inference = getAutoCommentText('inference', s.inference);
    if (!c.critical) c.critical = getAutoCommentText('critical', s.critical);
    if (!c.meta) c.meta = getAutoCommentText('meta', s.meta !== undefined ? s.meta : 60);

    // C. 비어있는 종합 소견 처방전 자동 구성
    if (!c.overall) {
        c.overall = getAutoOverallComment(student.name, s);
    }
}

/**
 * 6등급 세부 영역별 학술적 명품 코멘트 데이터셋
 */
const AUTO_COMMENTS = {
    vocab: [
        { min: 90, text: "어휘 인지 및 문맥 속 개념 추론 능력이 매우 독보적입니다. 비문학 지문에 자주 등장하는 학술적 개념어와 한자어의 정확한 함의를 직관적으로 이해하며, 풍부하고 정교한 어휘체계를 지니고 있습니다." },
        { min: 80, text: "문맥 속에서 낯선 어휘의 뜻을 유추하는 능력이 탁월하며, 한자어와 다의어의 의미를 다각도로 정확하게 이해하고 있습니다. 고난도 비문학 지문에 등장하는 개념어도 무리 없이 수용합니다." },
        { min: 70, text: "일상적인 학습 어휘의 의미를 잘 알고 있으며, 문장의 전후 맥락을 파악하여 생소한 어휘의 뜻을 적절히 유추해냅니다. 다만 교과 전용 고난도 한자어는 지속적인 어휘 정리가 필요합니다." },
        { min: 60, text: "기본적인 어휘력은 갖추고 있으나, 추상적인 개념어나 한자어의 정확한 사전적 정의 파악에 다소 어려움을 겪습니다. 단어를 문맥상 대략적으로만 이해하고 넘어가는 습관을 보완해야 합니다." },
        { min: 50, text: "글을 읽을 때 핵심 단어의 뜻을 잘 몰라 문장 전체의 의미를 잘못 이해하는 경우가 자주 발생합니다. 교과서 핵심 어휘와 기초 독해 한자어를 중심으로 매일 단어장을 작성하며 기초 체력을 길러야 합니다." },
        { min: 0,  text: "독해의 첫 단추인 단어 인지 단계에서 장벽을 느끼고 있습니다. 지문에 자주 빈출되는 일상 어휘와 교과 용어의 뜻을 소리 내어 읽고 뜻을 확인하는 1:1 집중 기초 어휘 학습이 반드시 선행되어야 합니다." }
    ],
    analysis: [
        { min: 90, text: "글의 거시적 구조와 미시적 정보 간의 유기적 연계를 완벽하게 입체적으로 파악합니다. 단락 간의 논리적 흐름(인과, 대조 등)을 도식화하는 구조 분석 능력이 매우 독보적이며 명료하게 요약해 냅니다." },
        { min: 80, text: "글의 구조를 한눈에 파악하고 문단 간의 논리적 연계 관계를 체계적으로 도식화하는 능력이 매우 뛰어납니다. 핵심어와 뒷받침 문장을 정확히 구분하여 글을 구조적으로 분석합니다." },
        { min: 70, text: "전체적인 주제와 중심 문장은 비교적 잘 찾아내지만, 정보량이 많은 긴 지문에서 문장 간의 유기적 상관관계나 세부 정보를 꼼꼼하게 도식화하는 과정은 추가적인 연습이 필요합니다." },
        { min: 60, text: "글을 읽고 요약하거나 대략적인 내용을 설명할 수는 있으나, 구조적인 관점(인과, 대조 등)에서 단락을 쪼개어 핵심 정보를 계통적으로 정리하는 힘이 부족합니다. 구조도 그리기 학습이 권장됩니다." },
        { min: 50, text: "단락별 중심 내용과 지엽적인 세부 사항을 구분하지 못하고 중심 흐름을 놓치는 편입니다. 문장을 읽을 때 핵심 문장과 수식 문장을 시각적으로 구분하여 표시하는 기초 분석 연습이 시급합니다." },
        { min: 0,  text: "정보량이 조금만 늘어나도 글의 전개 방향을 잃어버리는 상태입니다. 짧은 세 문장 분석에서 시작하여 주어와 서술어 찾기, 단락의 핵심어 밑줄 치기 등 기초 문장 분석부터 단계적으로 지도해야 합니다." }
    ],
    inference: [
        { min: 90, text: "텍스트 너머에 숨겨진 글쓴이의 전제, 암묵적 의도, 비유적 함의를 매우 정밀하고 깊이 있게 통찰해 냅니다. 고차원적 논리 추증 능력이 완전히 체화되어 낯선 소재의 고난도 문맥도 막힘없이 추론합니다." },
        { min: 80, text: "지문에 직접 드러나지 않은 글쓴이의 숨겨진 가정이나 전제를 매우 정교하게 논리적으로 추론해 냅니다. 맥락을 읽어내는 직관과 논리적 추증 능력이 완벽히 조화를 이룹니다." },
        { min: 70, text: "상식과 맥락에 기반한 합리적 추론 능력을 고루 갖추고 있어 행간의 의미를 무난하게 읽어냅니다. 복잡한 논리 구조를 가진 지문에서 전제를 논리적으로 도출하는 논리 훈련을 추가하면 완벽합니다." },
        { min: 60, text: "사실적으로 표기된 정보에만 의존하는 경향이 있어, 문장 뒤에 숨겨진 글쓴이의 의도나 비유적인 의미를 깊이 있게 파악하는 데는 한계가 있습니다. '왜 이런 표현을 썼을까?' 고민하는 깊이 있는 독서가 필요합니다." },
        { min: 50, text: "문장에 드러난 표면적 묘사에만 매몰되어, 생소한 문맥이 등장했을 때 유연하게 배경지식과 단서를 결합하여 뜻을 유추하지 못합니다. 단서를 모으고 논리적 가설을 세워보는 기초 질문식 독서가 필수적입니다." },
        { min: 0,  text: "문맥 속 단서들을 결합해 결론을 유추하는 것에 극심한 난조를 보입니다. '만약 그렇다면 다음엔 어떻게 될까?'와 같은 흥미 위주의 상황 가정 질문을 던져 사고의 확장을 이끌어내는 기초적 자극이 요구됩니다." }
    ],
    critical: [
        { min: 90, text: "저자의 입론과 논증 과정에서의 논리적 결함이나 숨은 편향성을 매우 객관적이고 날카롭게 변별해 냅니다. 다양한 관점의 비교 검토 능력이 뛰어나며 논리정연하고 입체적인 반론을 능동적으로 제안할 수 있습니다." },
        { min: 80, text: "저자의 주장과 근거 사이의 논리적 오류를 예리하게 짚어내고, 제시된 논조의 타당성과 객관성을 공정하게 검토하는 눈이 아주 매섭습니다. 자신만의 합리적인 대안이나 반론을 설득력 있게 제시할 수 있습니다." },
        { min: 70, text: "글에 드러난 정보의 객관성이나 편향성을 비교적 공정하게 선별할 수 있습니다. 단, 복잡하고 교묘한 주장에 내포된 이해관계나 숨은 의도까지 입체적으로 논박하는 깊이 있는 비판 능력을 좀 더 연마해야 합니다." },
        { min: 60, text: "권위 있는 저자의 주장이나 인쇄 매체에 담긴 정보를 별다른 의심 없이 그대로 수용하는 성향이 강합니다. 글의 논리적 허점을 찾거나 반대 관점의 근거를 스스로 구성해보는 다각적 사고 훈련이 유익합니다." },
        { min: 50, text: "주장과 사실을 구분하는 것을 어려워하며 글의 내용을 비판 없이 무조건 수용하는 독서 태도를 보입니다. 다양한 관점의 글을 비교 분석하여 읽고 '사실과 의견'을 명확히 구분하는 연습이 선행되어야 합니다." },
        { min: 0,  text: "글 속의 주장을 무조건적 진리로 무비판적으로 수용하는 독서 습관이 굳어져 있습니다. 일상 속 간단한 안건에 대해 찬반 의견을 직접 써보고, '이 글쓴이의 주장이 과연 정말 맞을까?' 의심해보는 기초 생각 열기가 필수적입니다." }
    ],
    meta: [
        { min: 90, text: "자신의 인지 과정을 완벽하게 모니터링하고 학습 전반을 주도적으로 조율 및 통제하는 능력이 압도적입니다. 모르는 내용과 아는 내용을 정확히 판별하여 전략적으로 지식을 수용하는 최고의 상위 인지력을 지니고 있습니다." },
        { min: 80, text: "자신의 독해 과정을 스스로 모니터링하고 부족한 부분을 찾아 보완하는 능력이 매우 안정적이며, 전략적 읽기를 잘 수행하고 있습니다. 문제 해결 시 오답의 원인을 논리적으로 짚어내며 주도적으로 보완합니다." },
        { min: 70, text: "기본적인 모니터링 습관이 잡혀 있으나, 긴 지문을 읽을 때 집중력이 떨어지면 자신의 이해 상태를 성찰하는 피드백 과정을 간과하곤 합니다. 이해도를 중간중간 자문하며 읽어가는 자기점검 습관이 요망됩니다." },
        { min: 60, text: "문제 풀이 후 오답 원인을 찾거나 모르는 어휘를 적극적으로 탐색하는 등의 자기조절 독서 훈련이 다소 부족합니다. '책을 읽으며 아는 내용과 모르는 내용'을 노트에 분류해보는 메타인지 연습이 효과적입니다." },
        { min: 50, text: "글을 기계적으로 읽는 성향이 강해, 자신이 내용을 제대로 파악하고 있는지에 대한 인지적 모니터링이 원활하지 않습니다. 문단이 끝날 때마다 짧은 질문을 통해 자신의 이해 여부를 매 순간 성찰해야 합니다." },
        { min: 0,  text: "자신의 지식 획득 과정을 객관화하는 연습이 전혀 되어 있지 않습니다. 학습에 대한 거부감을 줄이고, '내가 알고 있는 것'을 소리 내어 말해 보는 구술식 학습을 통해 메타인지의 첫걸음을 떼어야 합니다." }
    ]
};

/**
 * 특정 영역 및 점수에 부합하는 자동완성 소견 텍스트 반환
 */
function getAutoCommentText(domainKey, score) {
    const list = AUTO_COMMENTS[domainKey];
    if (!list) return "";
    for (let i = 0; i < list.length; i++) {
        if (score >= list[i].min) {
            return list[i].text;
        }
    }
    return "";
}

/**
 * 5대 점수를 믹싱하여 실시간 종합 소견 및 보완 학습 솔루션 처방전 자동 구성
 */
function getAutoOverallComment(name, s) {
    const domains = [
        { key: '어휘력', val: s.vocab !== undefined ? s.vocab : 70 },
        { key: '분석력', val: s.analysis !== undefined ? s.analysis : 70 },
        { key: '추론력', val: s.inference !== undefined ? s.inference : 70 },
        { key: '비판력', val: s.critical !== undefined ? s.critical : 70 },
        { key: '메타인지', val: s.meta !== undefined ? s.meta : 60 }
    ];

    // 내림차순 정렬하여 최대/최소값 획득
    domains.sort((a, b) => b.val - a.val);
    const strength = domains[0];
    const weakness = domains[domains.length - 1];

    return `${name} 학생은 영역 평가 항목 중 [${strength.key}] 영역에서 최고 성취도(${strength.val}점)를 보여주며, 문장의 핵심 의도를 날카롭게 파악하는 훌륭한 문해력 역량을 증명하였습니다. 다만, 상대적으로 주의 깊은 케어가 필요한 [${weakness.key}] 영역(${weakness.val}점)에 대해서는 세밀한 개별 교정이 필요합니다. 향후 [${weakness.key}] 영역을 집중적으로 연마할 수 있는 독해 훈련서 및 맞춤형 오답 클리닉 독서 훈련을 병행해 준다면, 전반적인 사고력이 단단하게 균형을 갖춘 최고 지성의 독서 리더로 성장해 나갈 수 있을 것입니다.`;
}
