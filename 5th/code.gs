/**
 * 구글 문서 LaTeX 수식 변환 도구
 * 
 * 작성자: 수학 교사용 도구
 * 버전: 1.0
 * 
 * ======================== 사용 설명서 ========================
 * 
 * 【설치 방법】
 * 1. 구글 문서 열기
 * 2. 확장프로그램 → Apps Script 클릭
 * 3. 기존 코드 삭제 후 이 코드 전체 붙여넣기
 * 4. 저장 (Ctrl+S)
 * 5. 실행 버튼 ▶ 클릭 → "onOpen" 선택 → 실행
 * 6. 권한 승인 (처음 한 번만)
 * 7. 구글 문서로 돌아가서 새로고침 (F5)
 * 
 * 【수식 입력 방법】
 * - 수식을 달러 기호($)로 감싸기
 * - 예: $\frac{1}{2}$ → 분수로 변환
 * 
 * 【지원되는 수식】
 * • 분수: $\frac{a}{b}$
 * • 극한: $\lim_{x \to 0}$
 * • 적분: $\int_0^1 x^2 dx$
 * • 미분: $\frac{d}{dx}$
 * • 제곱근: $\sqrt{x}$
 * • 위첨자: $x^2$
 * • 아래첨자: $x_n$
 * • 합: $\sum_{i=1}^n$
 * • 행렬: $\begin{pmatrix} a & b \\ c & d \end{pmatrix}$
 * • 집합: $A \cup B$, $A \cap B$, $x \in A$
 * • 그리스 문자: $\alpha$, $\beta$, $\gamma$, $\pi$
 * 
 * 【메뉴 설명】
 * - 수식 변환: 문서 전체의 수식을 이미지로 변환
 * - 선택 영역 변환: 선택한 부분만 변환
 * - 수동 변환: 팝업창에서 수식 입력
 * - 크기 설정: 학년에 맞게 이미지 크기 조절
 * - 사용법 보기: 이 도움말 표시
 * 
 * 【문제 해결】
 * - 변환 실패 시 "API 테스트" 실행
 * - 특수 문자 오류 시 더 간단한 수식으로 시도
 * - 네트워크 차단 시 다른 와이파이 사용
 * 
 * ============================================================
 */

// 전역 설정 저장소
var SETTINGS = PropertiesService.getDocumentProperties();

/**
 * 문서 열 때 메뉴 생성
 */
function onOpen() {
  const ui = DocumentApp.getUi();
  
  ui.createMenu('🔢 수식 변환')
    .addItem('✨ 수식 변환', 'convertAllFormulas')
    .addItem('📝 선택 영역 변환', 'convertSelection')
    .addItem('🔧 수동 변환', 'manualConvert')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 이미지 크기')
      .addItem('📏 작게 (고등학교)', 'setSizeSmall')
      .addItem('📐 보통 (중학교)', 'setSizeMedium')
      .addItem('📊 크게 (초등학교)', 'setSizeLarge')
      .addItem('🔍 매우 크게 (발표용)', 'setSizeXLarge')
      .addItem('❓ 현재 크기 확인', 'showCurrentSize'))
    .addSeparator()
    .addItem('🧪 API 테스트', 'testAPI')
    .addItem('📖 사용법 보기', 'showHelp')
    .addToUi();
    
  // 기본 크기 설정
  if (!SETTINGS.getProperty('imageSize')) {
    SETTINGS.setProperty('imageSize', '150');
  }
}

/**
 * 전체 문서 수식 변환
 */
function convertAllFormulas() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const ui = DocumentApp.getUi();
  
  let successCount = 0;
  let failCount = 0;
  let errors = [];
  
  const numChildren = body.getNumChildren();
  
  // 모든 단락 처리
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = child.asParagraph();
      const text = paragraph.getText();
      
      if (!text || !text.includes('$')) continue;
      
      // 수식 찾기 및 분리
      const parts = parseTextWithFormulas(text);
      
      if (parts.formulas.length > 0) {
        try {
          // 단락 재구성
          paragraph.clear();
          
          parts.segments.forEach(segment => {
            if (segment.type === 'text') {
              paragraph.appendText(segment.content);
            } else if (segment.type === 'formula') {
              const imageBlob = createFormulaImage(segment.content);
              
              if (imageBlob) {
                paragraph.appendInlineImage(imageBlob);
                successCount++;
              } else {
                // 실패한 수식은 텍스트로 표시
                paragraph.appendText('[수식: ' + segment.content + ']');
                failCount++;
                errors.push(segment.content);
              }
            }
          });
          
        } catch (e) {
          console.error('단락 처리 오류:', e.toString());
          failCount++;
        }
      }
    }
  }
  
  // 결과 메시지
  showResult(ui, successCount, failCount, errors);
}

/**
 * 선택 영역만 변환
 */
function convertSelection() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  const ui = DocumentApp.getUi();
  
  if (!selection) {
    ui.alert('알림', '변환할 텍스트를 먼저 선택해주세요.', ui.ButtonSet.OK);
    return;
  }
  
  let successCount = 0;
  let failCount = 0;
  
  const elements = selection.getSelectedElements();
  
  elements.forEach(element => {
    const elem = element.getElement();
    
    if (elem.getType() === DocumentApp.ElementType.TEXT || 
        elem.getType() === DocumentApp.ElementType.PARAGRAPH) {
      
      let paragraph;
      if (elem.getType() === DocumentApp.ElementType.TEXT) {
        paragraph = elem.getParent().asParagraph();
      } else {
        paragraph = elem.asParagraph();
      }
      
      const text = paragraph.getText();
      const parts = parseTextWithFormulas(text);
      
      if (parts.formulas.length > 0) {
        try {
          paragraph.clear();
          
          parts.segments.forEach(segment => {
            if (segment.type === 'text') {
              paragraph.appendText(segment.content);
            } else if (segment.type === 'formula') {
              const imageBlob = createFormulaImage(segment.content);
              
              if (imageBlob) {
                paragraph.appendInlineImage(imageBlob);
                successCount++;
              } else {
                paragraph.appendText('[수식: ' + segment.content + ']');
                failCount++;
              }
            }
          });
        } catch (e) {
          console.error('오류:', e.toString());
          failCount++;
        }
      }
    }
  });
  
  showResult(ui, successCount, failCount, []);
}

/**
 * 수동 변환 (팝업창)
 */
function manualConvert() {
  const ui = DocumentApp.getUi();
  
  const response = ui.prompt(
    '수동 수식 변환',
    'LaTeX 수식을 입력하세요 ($ 기호 없이):\n\n' +
    '예시:\n' +
    '• 분수: \\frac{1}{2}\n' +
    '• 극한: \\lim_{x \\to 0}\n' +
    '• 적분: \\int_0^1 x^2 dx\n' +
    '• 미분: \\frac{d}{dx}f(x)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const latex = response.getResponseText();
    
    if (latex && latex.trim()) {
      try {
        const imageBlob = createFormulaImage(latex.trim());
        
        if (imageBlob) {
          const doc = DocumentApp.getActiveDocument();
          const body = doc.getBody();
          body.appendParagraph('').appendInlineImage(imageBlob);
          ui.alert('성공', '수식이 삽입되었습니다!', ui.ButtonSet.OK);
        } else {
          ui.alert('실패', '이미지 생성에 실패했습니다.', ui.ButtonSet.OK);
        }
      } catch (e) {
        ui.alert('오류', e.toString(), ui.ButtonSet.OK);
      }
    }
  }
}

/**
 * 텍스트에서 수식 파싱
 */
function parseTextWithFormulas(text) {
  const segments = [];
  const formulas = [];
  
  if (!text || !text.includes('$')) {
    return { segments: [{type: 'text', content: text || ''}], formulas: [] };
  }
  
  let currentPos = 0;
  
  while (currentPos < text.length) {
    const startDollar = text.indexOf('$', currentPos);
    
    if (startDollar === -1) {
      // 남은 텍스트 추가
      if (currentPos < text.length) {
        segments.push({
          type: 'text',
          content: text.substring(currentPos)
        });
      }
      break;
    }
    
    // $ 앞의 텍스트 추가
    if (startDollar > currentPos) {
      segments.push({
        type: 'text',
        content: text.substring(currentPos, startDollar)
      });
    }
    
    // 수식 끝 찾기
    const endDollar = text.indexOf('$', startDollar + 1);
    
    if (endDollar === -1) {
      // 짝이 없는 $
      segments.push({
        type: 'text',
        content: text.substring(startDollar)
      });
      break;
    }
    
    // 수식 추출
    const formula = text.substring(startDollar + 1, endDollar).trim();
    if (formula) {
      segments.push({
        type: 'formula',
        content: formula
      });
      formulas.push(formula);
    }
    
    currentPos = endDollar + 1;
  }
  
  return { segments: segments, formulas: formulas };
}

/**
 * LaTeX 수식을 이미지로 변환
 */
function createFormulaImage(latex) {
  try {
    const size = SETTINGS.getProperty('imageSize') || '150';
    
    // CodeCogs API URL
    const baseUrl = 'https://latex.codecogs.com/png.latex?';
    const dpiParam = '\\dpi{' + size + '}';
    const fullUrl = baseUrl + encodeURIComponent(dpiParam + ' ' + latex);
    
    // 이미지 가져오기
    const response = UrlFetchApp.fetch(fullUrl, {
      'muteHttpExceptions': true,
      'followRedirects': true
    });
    
    if (response.getResponseCode() === 200) {
      const blob = response.getBlob();
      // 최소 크기 확인 (빈 이미지 방지)
      if (blob.getBytes().length > 100) {
        return blob;
      }
    }
    
    // 실패 시 DPI 없이 재시도
    const simpleUrl = baseUrl + encodeURIComponent(latex);
    const response2 = UrlFetchApp.fetch(simpleUrl, {
      'muteHttpExceptions': true
    });
    
    if (response2.getResponseCode() === 200) {
      return response2.getBlob();
    }
    
    return null;
    
  } catch (e) {
    console.error('이미지 생성 실패:', latex, e.toString());
    return null;
  }
}

/**
 * 크기 설정 함수들
 */
function setSizeSmall() {
  SETTINGS.setProperty('imageSize', '100');
  DocumentApp.getUi().alert('설정', '크기: 작게 (고등학교용)\nDPI: 100', DocumentApp.getUi().ButtonSet.OK);
}

function setSizeMedium() {
  SETTINGS.setProperty('imageSize', '150');
  DocumentApp.getUi().alert('설정', '크기: 보통 (중학교용)\nDPI: 150', DocumentApp.getUi().ButtonSet.OK);
}

function setSizeLarge() {
  SETTINGS.setProperty('imageSize', '200');
  DocumentApp.getUi().alert('설정', '크기: 크게 (초등학교용)\nDPI: 200', DocumentApp.getUi().ButtonSet.OK);
}

function setSizeXLarge() {
  SETTINGS.setProperty('imageSize', '250');
  DocumentApp.getUi().alert('설정', '크기: 매우 크게 (발표용)\nDPI: 250', DocumentApp.getUi().ButtonSet.OK);
}

function showCurrentSize() {
  const size = SETTINGS.getProperty('imageSize') || '150';
  const sizeNames = {
    '100': '작게 (고등학교)',
    '150': '보통 (중학교)',
    '200': '크게 (초등학교)',
    '250': '매우 크게 (발표용)'
  };
  
  DocumentApp.getUi().alert(
    '현재 설정',
    '크기: ' + (sizeNames[size] || '사용자 지정') + '\nDPI: ' + size,
    DocumentApp.getUi().ButtonSet.OK
  );
}

/**
 * API 테스트
 */
function testAPI() {
  const ui = DocumentApp.getUi();
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  const testCases = [
    { latex: '\\frac{1}{2}', name: '분수' },
    { latex: '\\int_0^1 x^2 dx', name: '적분' },
    { latex: '\\lim_{x \\to 0} \\frac{\\sin x}{x}', name: '극한' }
  ];
  
  let results = '=== API 테스트 결과 ===\n\n';
  let successCount = 0;
  
  body.appendParagraph('--- API 테스트 ---');
  
  testCases.forEach(test => {
    try {
      const imageBlob = createFormulaImage(test.latex);
      
      if (imageBlob) {
        results += `✅ ${test.name}: 성공\n`;
        body.appendParagraph(test.name + ': ').appendInlineImage(imageBlob);
        successCount++;
      } else {
        results += `❌ ${test.name}: 실패\n`;
      }
    } catch (e) {
      results += `❌ ${test.name}: 오류 - ${e.toString()}\n`;
    }
  });
  
  results += `\n총 ${successCount}/3 성공`;
  ui.alert('테스트 결과', results, ui.ButtonSet.OK);
}

/**
 * 사용법 표시
 */
function showHelp() {
  const ui = DocumentApp.getUi();
  
  const helpText = `
📖 수식 변환 도구 사용법

【수식 입력 방법】
• 수식을 $ 기호로 감싸기
• 예: $\\frac{1}{2}$ → 분수 이미지로 변환

【지원 수식 예시】
• 분수: $\\frac{a}{b}$
• 극한: $\\lim_{x \\to 0}$
• 적분: $\\int_0^1 x^2 dx$
• 미분: $\\frac{d}{dx}$
• 제곱근: $\\sqrt{x}$
• 집합: $A \\cup B$

【크기 조절】
• 초등학교: 크게 (200 DPI)
• 중학교: 보통 (150 DPI)
• 고등학교: 작게 (100 DPI)

【문제 해결】
1. API 테스트 실행으로 연결 확인
2. 더 간단한 수식으로 시도
3. 수동 변환 사용

【팁】
• 한 번에 여러 수식 변환 가능
• 실패한 수식은 [수식: ...] 형태로 표시
• 크기 설정은 문서별로 저장됨

문의: 구글 "LaTeX 수식" 검색
`;
  
  ui.alert('사용법', helpText, ui.ButtonSet.OK);
}

/**
 * 결과 표시 헬퍼 함수
 */
function showResult(ui, successCount, failCount, errors) {
  let message = '';
  
  if (successCount > 0) {
    message += `✅ 성공: ${successCount}개\n`;
  }
  
  if (failCount > 0) {
    message += `⚠️ 실패: ${failCount}개\n`;
    
    if (errors.length > 0 && errors.length <= 3) {
      message += '\n실패한 수식:\n';
      errors.forEach(err => {
        message += `• ${err}\n`;
      });
    }
  }
  
  if (successCount === 0 && failCount === 0) {
    message = '변환할 수식을 찾을 수 없습니다.\n$수식$ 형태로 입력했는지 확인하세요.';
  }
  
  ui.alert('변환 결과', message, ui.ButtonSet.OK);
}

/**
 * 디버그용 빠른 테스트
 */
function quickTest() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  body.appendParagraph('=== 빠른 테스트 ===');
  body.appendParagraph('분수: $\\frac{1}{2}$');
  body.appendParagraph('적분: $\\int_0^1 x^2 dx$');
  body.appendParagraph('극한: $\\lim_{x \\to 0} f(x)$');
  
  convertAllFormulas();
}
