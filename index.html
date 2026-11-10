<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>북돋움 독해력 분석 보고서 자동 생성 프로그램</title>
    <!-- Google Web Fonts: Noto Sans KR for Korean, Inter for numerals/English -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Noto+Sans+KR:wght@300;400;500;700;900&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="index.css">
    <!-- SheetJS Excel Parser Library -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
</head>
<body>
    <div class="app-container">
        <!-- 좌측 입력창 (컨트롤 패널) -->
        <aside class="input-pane">
            <header class="pane-header">
                <div class="logo-area">
                    <span class="logo-icon">📖</span>
                    <h1>북돋움 리포트 매니저</h1>
                </div>
                <p class="subtitle">학생별 독해 진단 보고서 실시간 관리</p>
            </header>

            <!-- 엑셀 데이터 불러오기 섹션 -->
            <section class="excel-import-section">
                <h2 class="section-title"><span class="number">📊</span> 엑셀 데이터 빠른 입력</h2>
                <div class="excel-dropzone" id="excel-dropzone">
                    <span class="dropzone-icon">📥</span>
                    <p class="dropzone-text">여기에 엑셀 파일을 끌어다 놓거나 클릭하세요.</p>
                    <span class="file-name" id="file-name-display">선택된 엑셀 파일 없음</span>
                    <input type="file" id="excel-file-input" accept=".xlsx, .xls" style="display: none;">
                </div>
                <!-- 학생 다중 선택 드롭다운 (파일 업로드 시에만 보임) -->
                <div class="student-select-group" id="student-select-group" style="display: none;">
                    <label for="excel-student-select">👤 엑셀 학생 데이터 선택</label>
                    <select id="excel-student-select" class="student-select-dropdown">
                        <!-- 동적으로 로드되는 학생들 옵션 -->
                    </select>
                </div>
            </section>

            <form class="control-form" onsubmit="event.preventDefault();">
                <!-- 학생 기본 정보 -->
                <section class="form-section">
                    <h2 class="section-title"><span class="number">1</span> 학생 기본 정보 입력</h2>
                    <div class="input-grid">
                        <div class="input-group">
                            <label for="student-name">학생 이름</label>
                            <input type="text" id="student-name" value="강이솔" placeholder="예: 홍길동">
                        </div>
                        <div class="input-group">
                            <label for="academy-name">학원 / 아동센터명</label>
                            <input type="text" id="academy-name" value="꿈동산지역아동센터" placeholder="예: 꿈나무학원">
                        </div>
                        <div class="input-group">
                            <label for="student-grade">학년</label>
                            <input type="text" id="student-grade" value="2026년 초4" placeholder="예: 2026년 초등학교 4학년">
                        </div>
                        <div class="input-group">
                            <label for="student-reading-level">독서 단계</label>
                            <select id="student-reading-level" class="reading-level-select">
                                <option value="1단계" selected>1단계</option>
                                <option value="2단계">2단계</option>
                                <option value="3단계">3단계</option>
                                <option value="4단계">4단계</option>
                                <option value="5단계">5단계</option>
                                <option value="6단계">6단계</option>
                            </select>
                        </div>
                        <div class="input-group">
                            <label for="student-reading-sublevel">세부 구분</label>
                            <select id="student-reading-sublevel" class="reading-level-select">
                                <option value="기본" selected>기본 (없음)</option>
                                <option value="디딤돌">디딤돌</option>
                                <option value="도약">도약</option>
                                <option value="비상">비상</option>
                            </select>
                        </div>
                    </div>
                </section>

                <!-- 독해력 4대 영역 점수 -->
                <section class="form-section">
                    <h2 class="section-title"><span class="number">2</span> 영역별 진단 점수 (0~100점)</h2>
                    <div class="scores-container">
                        <!-- 어휘력 -->
                        <div class="score-input-card">
                            <div class="score-meta">
                                <label for="score-vocab">어휘력</label>
                                <span class="badge" id="badge-vocab">상위권</span>
                            </div>
                            <div class="slider-row">
                                <input type="range" id="score-vocab" min="0" max="100" value="85" class="score-slider">
                                <input type="number" id="num-vocab" min="0" max="100" value="85" class="score-number">
                            </div>
                        </div>

                        <!-- 분석력 -->
                        <div class="score-input-card">
                            <div class="score-meta">
                                <label for="score-analysis">분석력</label>
                                <span class="badge" id="badge-analysis">평균 이상</span>
                            </div>
                            <div class="slider-row">
                                <input type="range" id="score-analysis" min="0" max="100" value="72" class="score-slider">
                                <input type="number" id="num-analysis" min="0" max="100" value="72" class="score-number">
                            </div>
                        </div>

                        <!-- 추론력 -->
                        <div class="score-input-card">
                            <div class="score-meta">
                                <label for="score-inference">추론력</label>
                                <span class="badge" id="badge-inference">평균</span>
                            </div>
                            <div class="slider-row">
                                <input type="range" id="score-inference" min="0" max="100" value="64" class="score-slider">
                                <input type="number" id="num-inference" min="0" max="100" value="64" class="score-number">
                            </div>
                        </div>

                        <!-- 비판력 -->
                        <div class="score-input-card">
                            <div class="score-meta">
                                <label for="score-critical">비판력</label>
                                <span class="badge" id="badge-critical">보완 필요</span>
                            </div>
                            <div class="slider-row">
                                <input type="range" id="score-critical" min="0" max="100" value="50" class="score-slider">
                                <input type="number" id="num-critical" min="0" max="100" value="50" class="score-number">
                            </div>
                        </div>
                    </div>
                </section>

                <!-- 세부 내용 코멘트 입력 -->
                <section class="form-section">
                    <h2 class="section-title"><span class="number">3</span> 영역별 세부 코멘트 입력</h2>
                    
                    <div class="textarea-group">
                        <label for="comment-vocab">어휘력 세부 의견</label>
                        <textarea id="comment-vocab" rows="2" placeholder="어휘력 영역에 대한 평가를 적어주세요.">교과서에 자주 등장하는 고난도 한자어와 필수 어휘의 뜻을 정확히 파악하고 있으며, 문맥을 활용한 유추 능력이 뛰어납니다.</textarea>
                    </div>

                    <div class="textarea-group">
                        <label for="comment-analysis">분석력 세부 의견</label>
                        <textarea id="comment-analysis" rows="2" placeholder="분석력 영역에 대한 평가를 적어주세요.">글의 핵심 구조를 파악하고 정보를 체계적으로 분류하는 능력이 우수하나, 긴 지문에서 핵심 문장 간의 연계 관계를 놓치는 경우가 있어 구조화 훈련이 필요합니다.</textarea>
                    </div>

                    <div class="textarea-group">
                        <label for="comment-inference">추론력 세부 의견</label>
                        <textarea id="comment-inference" rows="2" placeholder="추론력 영역에 대한 평가를 적어주세요.">지문에 명시되지 않은 정보를 논리적 근거에 기반하여 추론하는 과정에 깊이가 있으며, 숨겨진 주제를 파악하는 직관력이 돋보입니다.</textarea>
                    </div>

                    <div class="textarea-group">
                        <label for="comment-critical">비판력 세부 의견</label>
                        <textarea id="comment-critical" rows="2" placeholder="비판력 영역에 대한 평가를 적어주세요.">글쓴이의 관점과 의도를 객관적으로 검토하는 능력이 다소 부족하며, 주장의 타당성과 신뢰성을 비판적으로 판단하는 연습이 보완되어야 합니다.</textarea>
                    </div>
                </section>

                <!-- 종합 해석 입력 -->
                <section class="form-section">
                    <h2 class="section-title"><span class="number">4</span> 종합 해석 입력</h2>
                    <div class="textarea-group">
                        <label for="comment-overall">보고서 총평 및 추천 보완책</label>
                        <textarea id="comment-overall" rows="4" placeholder="학생을 위한 종합적인 해석과 공부 방향 제안을 작성해 주세요.">이솔 학생은 어휘력과 추론력 영역이 매우 우수하여 깊이 있는 독해가 가능하지만, 분석력과 비판적 사고력이 다소 약합니다. 앞으로 논설문이나 칼럼 등을 분석적으로 읽고, 글쓴이의 주장에 대한 반론을 스스로 제기해보는 비판적 글쓰기 훈련을 병행한다면 균형 잡힌 문해력을 완성할 수 있을 것입니다.</textarea>
                    </div>
                </section>

                <!-- 액션 버튼들 -->
                <div class="action-buttons">
                    <button type="button" id="btn-print" class="btn btn-primary">
                        <span class="icon">🖨️</span> 실시간 인쇄하기
                    </button>
                    <button type="button" id="btn-pdf" class="btn btn-secondary">
                        <span class="icon">💾</span> PDF로 저장하기
                    </button>
                </div>
            </form>
        </aside>

        <!-- 우측 결과창 (A4 크기 실시간 프리뷰) -->
        <main class="preview-pane">
            <div class="a4-container">
                <div class="a4-page" id="printable-report">
                    <!-- 보고서 헤더 -->
                    <div class="report-header">
                        <div class="logo-badge">북돋움</div>
                        <h1>독해력 진단 분석 보고서</h1>
                        <div class="header-divider"></div>
                    </div>

                    <!-- 학생 기본 정보 테이블 -->
                    <table class="student-info-table">
                        <tbody>
                            <tr>
                                <th>이름</th>
                                <td id="view-name">강이솔</td>
                                <th>학원 / 기관명</th>
                                <td id="view-center">꿈동산지역아동센터</td>
                            </tr>
                            <tr>
                                <th>학년</th>
                                <td id="view-grade">2026년 초4</td>
                                <th>독서 단계</th>
                                <td id="view-reading-level">1단계</td>
                            </tr>
                            <tr>
                                <th>진단 일자</th>
                                <td id="view-date" colspan="3">2026-05-22</td>
                            </tr>
                        </tbody>
                    </table>

                    <!-- 진단 결과 시각화 (그래프 & 수치 표) -->
                    <div class="diagnosis-results">
                        <!-- 방사형 거미줄 차트 (SVG 구현) -->
                        <div class="chart-container">
                            <h3 class="box-title">📊 독해 영역별 발달 프로필</h3>
                            <div class="svg-wrapper">
                                <svg viewBox="0 0 320 320" width="100%" height="100%" id="radar-chart">
                                    <defs>
                                        <linearGradient id="poly-grad" x1="0%" y1="0%" x2="100%" y2="100%">
                                            <stop offset="0%" stop-color="#3b82f6" stop-opacity="0.35" />
                                            <stop offset="100%" stop-color="#8b5cf6" stop-opacity="0.35" />
                                        </linearGradient>
                                    </defs>
                                    <!-- 4개 영역의 정사각형 방사형 차트 -->
                                    <polygon points="160,60 260,160 160,260 60,160" class="grid-line max" />
                                    <polygon points="160,80 240,160 160,240 80,160" class="grid-line" />
                                    <polygon points="160,100 220,160 160,220 100,160" class="grid-line" />
                                    <polygon points="160,120 200,160 160,200 120,160" class="grid-line" />
                                    <polygon points="160,140 180,160 160,180 140,160" class="grid-line" />
                                    <line x1="160" y1="160" x2="160" y2="60" class="axis-line" />
                                    <line x1="160" y1="160" x2="260" y2="160" class="axis-line" />
                                    <line x1="160" y1="160" x2="160" y2="260" class="axis-line" />
                                    <line x1="160" y1="160" x2="60" y2="160" class="axis-line" />
                                    <text x="164" y="84" class="grid-label">80</text>
                                    <text x="164" y="104" class="grid-label">60</text>
                                    <text x="164" y="124" class="grid-label">40</text>
                                    <text x="164" y="144" class="grid-label">20</text>
                                    <polygon points="" id="data-poly" class="data-polygon" />
                                    <circle cx="160" cy="160" r="4.5" id="dot-vocab" class="data-dot dot-vocab" />
                                    <circle cx="160" cy="160" r="4.5" id="dot-analysis" class="data-dot dot-analysis" />
                                    <circle cx="160" cy="160" r="4.5" id="dot-inference" class="data-dot dot-inference" />
                                    <circle cx="160" cy="160" r="4.5" id="dot-critical" class="data-dot dot-critical" />
                                    <text x="160" y="42" class="axis-label" text-anchor="middle">어휘력</text>
                                    <text x="270" y="164" class="axis-label" text-anchor="start">분석력</text>
                                    <text x="160" y="284" class="axis-label" text-anchor="middle">추론력</text>
                                    <text x="50" y="164" class="axis-label" text-anchor="end">비판력</text>
                                </svg>
                            </div>
                        </div>

                        <!-- 분석 결과 상세 수치 테이블 -->
                        <div class="result-table-container">
                            <h3 class="box-title">📈 평가 영역별 성취 현황</h3>
                            <table class="evaluation-table">
                                <thead>
                                    <tr>
                                        <th>평가 영역</th>
                                        <th>점수</th>
                                        <th>종합 평가</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr id="row-vocab">
                                        <td class="td-field">어휘력</td>
                                        <td class="td-score"><span id="val-vocab">85</span>점</td>
                                        <td class="td-rating"><span class="badge" id="rate-vocab">상위권</span></td>
                                    </tr>
                                    <tr id="row-analysis">
                                        <td class="td-field">분석력</td>
                                        <td class="td-score"><span id="val-analysis">72</span>점</td>
                                        <td class="td-rating"><span class="badge" id="rate-analysis">평균 이상</span></td>
                                    </tr>
                                    <tr id="row-inference">
                                        <td class="td-field">추론력</td>
                                        <td class="td-score"><span id="val-inference">64</span>점</td>
                                        <td class="td-rating"><span class="badge" id="rate-inference">평균</span></td>
                                    </tr>
                                    <tr id="row-critical">
                                        <td class="td-field">비판력</td>
                                        <td class="td-score"><span id="val-critical">50</span>점</td>
                                        <td class="td-rating"><span class="badge" id="rate-critical">보완 필요</span></td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- 영역별 세부 의견 기술 영역 -->
                    <div class="detail-comments-section">
                        <h3 class="box-title">📝 영역별 세부 진단 결과</h3>
                        <table class="comments-table">
                            <colgroup>
                                <col style="width: 22%">
                                <col style="width: 78%">
                            </colgroup>
                            <tbody>
                                <tr>
                                    <th class="col-header vocab-header">어휘력</th>
                                    <td id="view-vocab-comment">교과서에 자주 등장하는 고난도 한자어와 필수 어휘의 뜻을 정확히 파악하고 있으며, 문맥을 활용한 유추 능력이 뛰어납니다.</td>
                                </tr>
                                <tr>
                                    <th class="col-header analysis-header">분석력</th>
                                    <td id="view-analysis-comment">글의 핵심 구조를 파악하고 정보를 체계적으로 분류하는 능력이 우수하나, 긴 지문에서 핵심 문장 간의 연계 관계를 놓치는 경우가 있어 구조화 훈련이 필요합니다.</td>
                                </tr>
                                <tr>
                                    <th class="col-header inference-header">추론력</th>
                                    <td id="view-inference-comment">지문에 명시되지 않은 정보를 논리적 근거에 기반하여 추론하는 과정에 깊이가 있으며, 숨겨진 주제를 파악하는 직관력이 돋보입니다.</td>
                                </tr>
                                <tr>
                                    <th class="col-header critical-header">비판력</th>
                                    <td id="view-critical-comment">글쓴이의 관점과 의도를 객관적으로 검토하는 능력이 다소 부족하며, 주장의 타당성과 신뢰성을 비판적으로 판단하는 연습이 보완되어야 합니다.</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                    <!-- 종합 해석 및 처방 -->
                    <div class="overall-prescription-section">
                        <h3 class="box-title">💡 종합 소견 및 보완 학습 제언</h3>
                        <div class="prescription-box">
                            <p id="view-overall-comment">이솔 학생은 어휘력과 추론력 영역이 매우 우수하여 깊이 있는 독해가 가능하지만, 분석력과 비판적 사고력이 다소 약합니다. 앞으로 논설문이나 칼럼 등을 분석적으로 읽고, 글쓴이의 주장에 대한 반론을 스스로 제기해보는 비판적 글쓰기 훈련을 병행한다면 균형 잡힌 문해력을 완성할 수 있을 것입니다.</p>
                        </div>
                    </div>

                    <!-- 보고서 발행 기관 하단부 -->
                    <footer class="report-footer">
                        <p class="footer-notice">본 평가지는 독해력 진단 표준 지표에 근거하여 작성된 신뢰도 높은 분석 리포트입니다.</p>
                        <div class="org-signature">
                            <div class="date-stamp" id="view-footer-date">2026년 05월 22일</div>
                            <div class="org-name">
                                <span>북돋움 문해력 교육 연구소</span>
                                <div class="seal-container">
                                    <svg viewBox="0 0 100 100" width="48" height="48" class="red-seal">
                                        <!-- 원형 공식 붉은 직인 -->
                                        <circle cx="50" cy="50" r="43" fill="none" stroke="#dc2626" stroke-width="2" stroke-dasharray="0" />
                                        <circle cx="50" cy="50" r="39" fill="none" stroke="#dc2626" stroke-width="1.2" />
                                        <!-- 한글 전각 서체 형상화 (북돋움 교육소) -->
                                        <rect x="25" y="25" width="50" height="50" fill="none" stroke="#dc2626" stroke-width="1.5" rx="3" />
                                        <text x="50" y="45" font-size="12" font-weight="900" fill="#dc2626" text-anchor="middle" font-family="Noto Sans KR">북돋</text>
                                        <text x="50" y="65" font-size="12" font-weight="900" fill="#dc2626" text-anchor="middle" font-family="Noto Sans KR">움인</text>
                                    </svg>
                                </div>
                            </div>
                        </div>
                    </footer>
                </div>
            </div>
        </main>
    </div>

    <!-- 앱 로직 및 데이터 컨트롤 바인딩 -->
    <script src="app.js"></script>
</body>
</html>
