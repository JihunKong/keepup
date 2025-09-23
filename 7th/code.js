function createSocialStudiesPackage() {
  // 메인 실행 함수
  try {
    console.log('통합사회 수업 자료 패키지 생성을 시작합니다...');
    
    // 1. Google Forms 생성
    const formUrl = createAssessmentForm();
    console.log('✓ Google Forms 생성 완료: ' + formUrl);
    
    // 2. Google Docs 생성
    const docUrl = createActivityWorksheet();
    console.log('✓ Google Docs 생성 완료: ' + docUrl);
    
    // 3. Google Slides 생성
    const slidesUrl = createPresentationSlides();
    console.log('✓ Google Slides 생성 완료: ' + slidesUrl);
    
    // 결과 요약
    const summary = `
    ========================================
    통합사회 수업 자료 패키지 생성 완료!
    ========================================
    
    1. 평가 설문지 (Forms): ${formUrl}
    2. 활동지 (Docs): ${docUrl}
    3. 수업 자료 (Slides): ${slidesUrl}
    
    모든 자료가 Google Drive에 저장되었습니다.
    ========================================
    `;
    
    console.log(summary);
    
    // 결과를 스프레드시트에 기록 (선택사항)
    recordCreatedResources(formUrl, docUrl, slidesUrl);
    
    return {
      success: true,
      formUrl: formUrl,
      docUrl: docUrl,
      slidesUrl: slidesUrl
    };
    
  } catch (error) {
    console.error('오류 발생:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// 1. Google Forms 생성 함수 (수정됨)
function createAssessmentForm() {
  const form = FormApp.create('[통합사회] 정보화 사회 이해도 점검');
  
  // ⭐ 중요: 퀴즈 모드를 먼저 활성화해야 함
  form.setIsQuiz(true);
  
  // 폼 설명 추가
  form.setDescription(
    '이 설문은 정보화 사회에 대한 여러분의 이해도를 점검하기 위한 것입니다.\n' +
    '학습 내용을 떠올리며 성실하게 답변해 주세요.'
  );
  
  // 학생 정보 입력
  form.addTextItem()
    .setTitle('학번과 이름')
    .setHelpText('예: 10101 홍길동')
    .setRequired(true);
  
  // 섹션 1: 정보화 개념 확인
  form.addPageBreakItem()
    .setTitle('Part 1. 정보화 개념 확인');
  
  // 객관식 문제 1
  const q1 = form.addMultipleChoiceItem();
  q1.setTitle('1. 정보화 사회의 가장 핵심적인 특징은 무엇입니까?')
    .setPoints(5)
    .setChoices([
      q1.createChoice('정보와 지식이 핵심 자원이 되는 사회', true),
      q1.createChoice('농업이 주요 산업인 사회', false),
      q1.createChoice('제조업 중심의 산업 사회', false),
      q1.createChoice('서비스업만 발달한 사회', false)
    ])
    .setRequired(true)
    .setFeedbackForCorrect(FormApp.createFeedback()
      .setText('정확합니다! 정보화 사회는 정보와 지식이 핵심 자원이 되는 사회입니다.')
      .build())
    .setFeedbackForIncorrect(FormApp.createFeedback()
      .setText('다시 생각해보세요. 정보화 사회의 핵심은 정보와 지식입니다.')
      .build());
  
  // 객관식 문제 2
  const q2 = form.addMultipleChoiceItem();
  q2.setTitle('2. 다음 중 정보화 사회의 긍정적 영향이 아닌 것은?')
    .setPoints(5)
    .setChoices([
      q2.createChoice('시공간 제약 극복', false),
      q2.createChoice('정보 접근성 향상', false),
      q2.createChoice('개인정보 유출 증가', true),
      q2.createChoice('의사소통 방식의 다양화', false)
    ])
    .setRequired(true)
    .setFeedbackForCorrect(FormApp.createFeedback()
      .setText('맞습니다! 개인정보 유출은 정보화의 부정적 영향입니다.')
      .build())
    .setFeedbackForIncorrect(FormApp.createFeedback()
      .setText('개인정보 유출은 정보화의 부정적 측면입니다.')
      .build());
  
  // 섹션 2: 디지털 격차 사례
  form.addPageBreakItem()
    .setTitle('Part 2. 디지털 격차 사례');
  
  // 객관식 문제 3
  const q3 = form.addMultipleChoiceItem();
  q3.setTitle('3. 디지털 격차(Digital Divide)란 무엇을 의미합니까?')
    .setPoints(5)
    .setChoices([
      q3.createChoice('디지털 기기의 가격 차이', false),
      q3.createChoice('정보 접근과 활용 능력의 불평등', true),
      q3.createChoice('인터넷 속도의 차이', false),
      q3.createChoice('스마트폰 보유율의 차이', false)
    ])
    .setRequired(true)
    .setFeedbackForCorrect(FormApp.createFeedback()
      .setText('정확합니다! 디지털 격차는 정보 접근과 활용 능력의 불평등을 의미합니다.')
      .build())
    .setFeedbackForIncorrect(FormApp.createFeedback()
      .setText('디지털 격차는 단순한 기기 보유가 아닌 정보 접근과 활용 능력의 차이입니다.')
      .build());
  
  // 객관식 문제 4
  const q4 = form.addMultipleChoiceItem();
  q4.setTitle('4. 다음 중 디지털 격차 해소를 위한 정책으로 적절한 것은?')
    .setPoints(5)
    .setChoices([
      q4.createChoice('인터넷 사용 제한', false),
      q4.createChoice('스마트폰 사용 금지', false),
      q4.createChoice('디지털 리터러시 교육 확대', true),
      q4.createChoice('온라인 서비스 축소', false)
    ])
    .setRequired(true)
    .setFeedbackForCorrect(FormApp.createFeedback()
      .setText('훌륭합니다! 디지털 리터러시 교육은 격차 해소의 핵심입니다.')
      .build())
    .setFeedbackForIncorrect(FormApp.createFeedback()
      .setText('디지털 격차 해소를 위해서는 교육과 지원이 필요합니다.')
      .build());
  
  // 섹션 3: 서술형 문제
  form.addPageBreakItem()
    .setTitle('Part 3. 해결방안 제시');
  
  // 서술형 문제
  const q5 = form.addParagraphTextItem();
  q5.setTitle('5. 우리 사회의 디지털 격차 문제를 해결하기 위한 방안을 2가지 이상 구체적으로 제시하시오.')
    .setHelpText('각 방안에 대해 3문장 이상으로 설명해 주세요.')
    .setPoints(10)
    .setRequired(true);
  
  // 일반 피드백 설정
  q5.setGeneralFeedback(FormApp.createFeedback()
    .setText('다양한 관점에서 해결방안을 제시했는지 확인해보세요. 정부, 기업, 개인 차원의 노력이 모두 필요합니다.')
    .build());
  
  // 폼 설정
  form.setCollectEmail(false)
    .setProgressBar(true)
    .setConfirmationMessage('응답이 제출되었습니다. 수고하셨습니다!')
    .setAllowResponseEdits(true)
    .setLimitOneResponsePerUser(false)
    .setShuffleQuestions(false);
  
  return form.getPublishedUrl();
}

// 2. Google Docs 생성 함수 (변경 없음)
function createActivityWorksheet() {
  const doc = DocumentApp.create('[통합사회] 정보격차 사례 조사 활동지');
  const body = doc.getBody();
  
  // 스타일 정의
  const titleStyle = {};
  titleStyle[DocumentApp.Attribute.FONT_SIZE] = 20;
  titleStyle[DocumentApp.Attribute.BOLD] = true;
  
  const headingStyle = {};
  headingStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
  headingStyle[DocumentApp.Attribute.BOLD] = true;
  
  const normalStyle = {};
  normalStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  
  // 제목
  const title = body.appendParagraph('[통합사회] 정보격차 사례 조사 활동지');
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  title.setAttributes(titleStyle);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  // 학생 정보
  body.appendParagraph('\n학년/반: _______  모둠: _______  작성일: __________')
    .setAttributes(normalStyle);
  
  body.appendParagraph('모둠원: ________________________________________________')
    .setAttributes(normalStyle);
  
  // 학습목표
  body.appendParagraph('\n1. 학습목표')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAttributes(headingStyle);
  
  const objectives = body.appendParagraph(
    '• 정보화에 따른 생활 공간과 생활 양식의 변화를 설명할 수 있다.\n' +
    '• 정보격차의 원인과 영향을 분석할 수 있다.\n' +
    '• 정보격차 해소를 위한 다양한 방안을 제시할 수 있다.'
  );
  objectives.setAttributes(normalStyle);
  
  // 성취기준
  body.appendParagraph('\n▣ 성취기준')
    .setAttributes(headingStyle);
  
  body.appendParagraph(
    '[10통사02-03] 교통·통신의 발달과 정보화로 인한 생활공간과 생활양식의 변화 양상을 조사하고, ' +
    '이에 따른 문제점을 해결하기 위한 방안을 제안한다.'
  ).setAttributes(normalStyle);
  
  // 모둠별 조사 활동 - 국내 사례
  body.appendParagraph('\n2. 모둠별 조사 활동')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAttributes(headingStyle);
  
  body.appendParagraph('\n가. 국내 정보격차 사례 조사')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAttributes(headingStyle);
  
  // 표 생성 - 국내 사례
  const table1 = body.appendTable();
  const headers1 = ['구분', '내용'];
  const rows1 = [
    ['대상 집단', '(예: 고령층, 농어촌 지역, 저소득층 등)'],
    ['격차 현황', '\n\n\n'],
    ['발생 원인', '\n\n\n'],
    ['사회적 영향', '\n\n\n'],
    ['해결 노력', '\n\n\n']
  ];
  
  // 헤더 추가
  const headerRow1 = table1.appendTableRow();
  headers1.forEach(header => {
    const cell = headerRow1.appendTableCell(header);
    cell.setBackgroundColor('#e8f0fe');
    cell.getChild(0).asParagraph().setAttributes(headingStyle);
  });
  
  // 데이터 행 추가
  rows1.forEach(row => {
    const tableRow = table1.appendTableRow();
    row.forEach((cell, index) => {
      const tableCell = tableRow.appendTableCell(cell);
      if (index === 0) {
        tableCell.setBackgroundColor('#f8f9fa');
        tableCell.getChild(0).asParagraph().setAttributes(headingStyle);
      } else {
        tableCell.getChild(0).asParagraph().setAttributes(normalStyle);
      }
    });
  });
  
  // 국외 사례
  body.appendParagraph('\n나. 국외 정보격차 사례 조사')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAttributes(headingStyle);
  
  // 표 생성 - 국외 사례
  const table2 = body.appendTable();
  const rows2 = [
    ['국가/지역', ''],
    ['격차 현황', '\n\n\n'],
    ['특징', '\n\n\n'],
    ['해결 정책', '\n\n\n'],
    ['시사점', '\n\n\n']
  ];
  
  // 헤더 추가
  const headerRow2 = table2.appendTableRow();
  headers1.forEach(header => {
    const cell = headerRow2.appendTableCell(header);
    cell.setBackgroundColor('#e8f0fe');
    cell.getChild(0).asParagraph().setAttributes(headingStyle);
  });
  
  // 데이터 행 추가
  rows2.forEach(row => {
    const tableRow = table2.appendTableRow();
    row.forEach((cell, index) => {
      const tableCell = tableRow.appendTableCell(cell);
      if (index === 0) {
        tableCell.setBackgroundColor('#f8f9fa');
        tableCell.getChild(0).asParagraph().setAttributes(headingStyle);
      } else {
        tableCell.getChild(0).asParagraph().setAttributes(normalStyle);
      }
    });
  });
  
  // 개인 성찰 및 해결방안
  body.appendPageBreak();
  
  body.appendParagraph('3. 개인 성찰 및 해결방안')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAttributes(headingStyle);
  
  body.appendParagraph('\n가. 나의 디지털 활용 수준 점검')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAttributes(headingStyle);
  
  body.appendParagraph(
    '□ 정보 검색 능력: ☆☆☆☆☆\n' +
    '□ 디지털 기기 활용: ☆☆☆☆☆\n' +
    '□ 온라인 소통 능력: ☆☆☆☆☆\n' +
    '□ 디지털 콘텐츠 제작: ☆☆☆☆☆\n' +
    '□ 정보 윤리 의식: ☆☆☆☆☆'
  ).setAttributes(normalStyle);
  
  body.appendParagraph('\n나. 우리가 제안하는 정보격차 해결방안')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAttributes(headingStyle);
  
  body.appendParagraph(
    '\n방안 1: ___________________________________________________________\n' +
    '구체적 실천 방법: _________________________________________________\n' +
    '________________________________________________________________\n\n' +
    '방안 2: ___________________________________________________________\n' +
    '구체적 실천 방법: _________________________________________________\n' +
    '________________________________________________________________\n\n' +
    '방안 3: ___________________________________________________________\n' +
    '구체적 실천 방법: _________________________________________________\n' +
    '________________________________________________________________'
  ).setAttributes(normalStyle);
  
  // 평가 기준
  body.appendParagraph('\n4. 평가 기준')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAttributes(headingStyle);
  
  const criteriaTable = body.appendTable();
  const criteriaHeaders = ['평가 항목', '배점', '자기평가'];
  const criteriaRows = [
    ['사례 조사의 충실성', '30점', ''],
    ['원인 분석의 타당성', '25점', ''],
    ['해결방안의 구체성', '25점', ''],
    ['모둠 협력도', '20점', ''],
    ['합계', '100점', '']
  ];
  
  // 평가표 헤더
  const critHeaderRow = criteriaTable.appendTableRow();
  criteriaHeaders.forEach(header => {
    const cell = critHeaderRow.appendTableCell(header);
    cell.setBackgroundColor('#fce4ec');
    cell.getChild(0).asParagraph().setAttributes(headingStyle);
  });
  
  // 평가표 데이터
  criteriaRows.forEach((row, rowIndex) => {
    const tableRow = criteriaTable.appendTableRow();
    row.forEach((cell, index) => {
      const tableCell = tableRow.appendTableCell(cell);
      if (rowIndex === criteriaRows.length - 1) {
        tableCell.setBackgroundColor('#f8f9fa');
        tableCell.getChild(0).asParagraph().setAttributes(headingStyle);
      } else if (index === 0) {
        tableCell.getChild(0).asParagraph().setAttributes(normalStyle);
      }
    });
  });
  
  return doc.getUrl();
}

// 3. Google Slides 생성 함수 (변경 없음)
function createPresentationSlides() {
  const presentation = SlidesApp.create('[통합사회] 교통·통신 발달과 생활 변화');
  const slides = presentation.getSlides();
  
  // 기본 슬라이드 삭제
  slides[0].remove();
  
  // 슬라이드 1: 표지
  const slide1 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE);
  slide1.getShapes()[0].getText().setText('교통·통신 발달과 생활 변화');
  slide1.getShapes()[1].getText().setText(
    '고등학교 1학년 통합사회\n' +
    '단원: 생활공간과 사회\n\n' +
    new Date().toLocaleDateString('ko-KR')
  );
  slide1.getBackground().setSolidFill('#1a73e8');
  slide1.getShapes()[0].getText().getTextStyle().setForegroundColor('#ffffff').setFontSize(36).setBold(true);
  slide1.getShapes()[1].getText().getTextStyle().setForegroundColor('#ffffff').setFontSize(18);
  
  // 슬라이드 2: 학습목표와 성취기준
  const slide2 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  slide2.getShapes()[0].getText().setText('학습목표 및 성취기준');
  slide2.getShapes()[1].getText().setText(
    '학습목표\n' +
    '• 교통·통신의 발달이 생활공간에 미친 영향을 설명할 수 있다\n' +
    '• 정보화로 인한 생활양식의 변화를 분석할 수 있다\n' +
    '• 정보격차 문제와 해결방안을 제시할 수 있다\n\n' +
    '성취기준\n' +
    '[10통사02-03] 교통·통신의 발달과 정보화로 인한 생활공간과 생활양식의 변화 양상을 조사하고, 이에 따른 문제점을 해결하기 위한 방안을 제안한다'
  );
  
  // 슬라이드 3: 교통 발달 사례 1
  const slide3 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_TWO_COLUMNS);
  slide3.getShapes()[0].getText().setText('교통 발달이 가져온 변화 (1)');
  slide3.getShapes()[1].getText().setText(
    '고속철도(KTX)의 영향\n\n' +
    '• 전국 반나절 생활권 실현\n' +
    '• 지역 간 교류 활성화\n' +
    '• 출퇴근 가능 지역 확대\n' +
    '• 수도권 집중 가속화'
  );
  slide3.getShapes()[2].getText().setText(
    '항공 교통의 발달\n\n' +
    '• 국제 교류 증가\n' +
    '• 관광 산업 성장\n' +
    '• 글로벌 비즈니스 확대\n' +
    '• 문화 교류 활성화'
  );
  
  // 슬라이드 4: 교통 발달 사례 2
  const slide4 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  slide4.getShapes()[0].getText().setText('교통 발달이 가져온 변화 (2)');
  slide4.getShapes()[1].getText().setText(
    '도시 교통 체계의 변화\n\n' +
    '🚇 대중교통 발달\n' +
    '• 지하철 노선 확대로 이동 시간 단축\n' +
    '• 버스 정보 시스템으로 대기 시간 감소\n\n' +
    '🚗 개인 이동수단 다양화\n' +
    '• 전기차, 자율주행차 등장\n' +
    '• 공유 모빌리티 서비스 확산\n\n' +
    '⚠️ 발생한 문제점\n' +
    '• 교통 혼잡과 환경 오염\n' +
    '• 교통 소외 지역 발생'
  );
  
  // 슬라이드 5: 통신 발달 사례 1
  const slide5 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_TWO_COLUMNS);
  slide5.getShapes()[0].getText().setText('통신 발달이 가져온 변화 (1)');
  slide5.getShapes()[1].getText().setText(
    '인터넷과 스마트폰\n\n' +
    '📱 일상생활의 변화\n' +
    '• 언제 어디서나 정보 접근\n' +
    '• 실시간 소통 가능\n' +
    '• 모바일 쇼핑/뱅킹\n' +
    '• 원격 교육/근무'
  );
  slide5.getShapes()[2].getText().setText(
    'SNS의 영향\n\n' +
    '🌐 소통 방식의 변화\n' +
    '• 전 세계와 연결\n' +
    '• 정보 공유 속도 증가\n' +
    '• 여론 형성 채널 다양화\n' +
    '• 개인 미디어 시대'
  );
  
  // 슬라이드 6: 통신 발달 사례 2
  const slide6 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  slide6.getShapes()[0].getText().setText('통신 발달이 가져온 변화 (2)');
  slide6.getShapes()[1].getText().setText(
    '5G와 사물인터넷(IoT) 시대\n\n' +
    '🔗 초연결 사회의 특징\n' +
    '• 스마트홈: 가전제품의 지능화\n' +
    '• 스마트시티: 도시 인프라 연결\n' +
    '• 원격의료: 의료 서비스 접근성 향상\n' +
    '• 메타버스: 가상과 현실의 융합\n\n' +
    '💡 긍정적 영향\n' +
    '• 생활의 편리성 증대\n' +
    '• 새로운 산업과 일자리 창출\n' +
    '• 시공간 제약 극복'
  );
  
  // 슬라이드 7: 정보격차 문제 1
  const slide7 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  slide7.getShapes()[0].getText().setText('정보격차(Digital Divide) 문제 (1)');
  slide7.getShapes()[1].getText().setText(
    '정보격차란?\n' +
    '정보 접근과 활용 능력의 차이로 인한 사회적 불평등\n\n' +
    '주요 격차 유형\n\n' +
    '1️⃣ 세대 간 격차\n' +
    '• 고령층의 디지털 기기 활용 어려움\n' +
    '• 온라인 서비스 이용 제한\n\n' +
    '2️⃣ 지역 간 격차\n' +
    '• 도시와 농어촌의 인프라 차이\n' +
    '• 통신 서비스 품질 격차\n\n' +
    '3️⃣ 소득 계층 간 격차\n' +
    '• 디지털 기기 구입 능력 차이\n' +
    '• 교육 기회 불평등'
  );
  
  // 슬라이드 8: 정보격차 문제 2
  const slide8 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_TWO_COLUMNS);
  slide8.getShapes()[0].getText().setText('정보격차(Digital Divide) 문제 (2)');
  slide8.getShapes()[1].getText().setText(
    '정보격차의 영향\n\n' +
    '📉 부정적 결과\n' +
    '• 교육 기회 불평등 심화\n' +
    '• 경제적 격차 확대\n' +
    '• 사회 참여 제한\n' +
    '• 민주주의 위협'
  );
  slide8.getShapes()[2].getText().setText(
    '실제 사례\n\n' +
    '🔍 코로나19 시기\n' +
    '• 온라인 수업 참여 격차\n' +
    '• 재택근무 가능 여부\n' +
    '• 백신 예약 시스템 접근성\n' +
    '• 비대면 서비스 이용 제한'
  );
  
  // 슬라이드 9: 해결방안 토론
  const slide9 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  slide9.getShapes()[0].getText().setText('정보격차 해결방안 토론');
  slide9.getShapes()[1].getText().setText(
    '🤝 함께 생각해 봅시다\n\n' +
    '정부 차원의 노력\n' +
    '• 디지털 인프라 구축 확대\n' +
    '• 정보화 교육 프로그램 운영\n' +
    '• 취약계층 기기 지원\n\n' +
    '기업의 사회적 책임\n' +
    '• 접근성 높은 서비스 개발\n' +
    '• 디지털 교육 지원\n' +
    '• 공공 와이파이 확대\n\n' +
    '개인과 공동체의 역할\n' +
    '• 디지털 재능 기부\n' +
    '• 세대 간 멘토링\n' +
    '• 정보 윤리 실천\n\n' +
    '💭 토론 주제: 우리 학교/지역에서 실천 가능한 정보격차 해소 방안은?'
  );
  
  // 슬라이드 10: 정리 및 차시 예고
  const slide10 = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  slide10.getShapes()[0].getText().setText('학습 정리 및 차시 예고');
  slide10.getShapes()[1].getText().setText(
    '오늘의 학습 정리\n\n' +
    '✅ 교통·통신 발달의 영향\n' +
    '• 생활공간의 확대와 시공간 압축\n' +
    '• 생활양식의 근본적 변화\n\n' +
    '✅ 정보화 사회의 명암\n' +
    '• 긍정: 편의성, 효율성, 연결성 증대\n' +
    '• 부정: 정보격차, 프라이버시, 중독 문제\n\n' +
    '✅ 더불어 사는 정보 사회\n' +
    '• 디지털 포용 정책의 필요성\n' +
    '• 모두가 함께하는 정보화\n\n' +
    '📚 차시 예고\n' +
    '다음 시간: 자연환경과 인간생활\n' +
    '준비물: 세계지도, 색연필'
  );
  
  // 모든 슬라이드 텍스트 스타일 통일
  const allSlides = presentation.getSlides();
  allSlides.forEach((slide, index) => {
    if (index > 0) {  // 표지 제외
      slide.getShapes().forEach((shape, shapeIndex) => {
        if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
          const textRange = shape.getText();
          if (shapeIndex === 0) {  // 제목
            textRange.getTextStyle()
              .setFontSize(32)
              .setBold(true)
              .setForegroundColor('#1a73e8');
          } else {  // 본문
            textRange.getTextStyle()
              .setFontSize(16)
              .setForegroundColor('#202124');
          }
        }
      });
    }
  });
  
  return presentation.getUrl();
}

// 4. 생성된 자료 기록 함수 (선택사항)
function recordCreatedResources(formUrl, docUrl, slidesUrl) {
  try {
    // 기록용 스프레드시트 생성 또는 열기
    let spreadsheet;
    const files = DriveApp.getFilesByName('[통합사회] 수업자료 관리대장');
    
    if (files.hasNext()) {
      spreadsheet = SpreadsheetApp.open(files.next());
    } else {
      spreadsheet = SpreadsheetApp.create('[통합사회] 수업자료 관리대장');
      const sheet = spreadsheet.getActiveSheet();
      
      // 헤더 설정
      sheet.getRange('A1:F1').setValues([
        ['생성일시', '단원명', 'Forms URL', 'Docs URL', 'Slides URL', '비고']
      ]);
      sheet.getRange('A1:F1').setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');
    }
    
    const sheet = spreadsheet.getActiveSheet();
    
    // 새 행 추가
    const newRow = [
      new Date(),
      '정보화로 인한 생활 변화',
      formUrl,
      docUrl,
      slidesUrl,
      '자동 생성됨'
    ];
    
    sheet.appendRow(newRow);
    
    // 열 너비 자동 조정
    sheet.autoResizeColumns(1, 6);
    
    console.log('자료 기록 완료: ' + spreadsheet.getUrl());
    
  } catch (error) {
    console.log('기록 생성 실패 (선택사항):', error);
  }
}

// 5. 개별 실행 함수들 (테스트용)
function testFormOnly() {
  const url = createAssessmentForm();
  console.log('Form URL:', url);
}

function testDocOnly() {
  const url = createActivityWorksheet();
  console.log('Doc URL:', url);
}

function testSlidesOnly() {
  const url = createPresentationSlides();
  console.log('Slides URL:', url);
}
