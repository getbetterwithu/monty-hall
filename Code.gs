/**
 * 몬티홀 시뮬레이터 — 구글 시트(모둠추첨 기록) 연동 백엔드
 * 
 * 스프레드시트 구조:
 *  - [전체기록] 탭: 저장일시 | 반 | 회차 | 모둠 | 번호 | 성명 | 학생ID | 결과ID | 좌석순서
 *  - [명렬표] 탭: 학생ID | 반 | 번호 | 성명 | 성적총점 | 재적 | 전입 | 수정일시
 *  - [몬티홀실험] 탭: (자동생성) 저장일시 | 반 | 번호 | 성명 | 모둠 | 역할 | 총시행 | 바꿈승률 | 유지승률
 */

var LOG_SHEET = '전체기록';
var ROSTER_SHEET = '명렬표';
var MONTY_SHEET = '몬티홀실험';

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  
  // 현재 접속한 구글 계정 확인 (학교 계정 @snu.ms.kr)
  var email = '';
  try {
    email = Session.getActiveUser().getEmail();
  } catch (err) {
    email = '';
  }
  
  template.userEmail = email;
  
  return template.evaluate()
    .setTitle('몬티홀 문제 확률 실험실')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 접속한 학생의 이메일 또는 학번을 기반으로 최신 좌석 및 모둠 짝꿍을 자동 조회
 */
function getStudentSeatAndPair(userEmail) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(LOG_SHEET);
    if (!logSheet || logSheet.getLastRow() < 2) {
      return { ok: false, error: '좌석 추첨 기록이 아직 없습니다.' };
    }

    var logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 9).getValues();
    
    // 1. 최신 회차 찾기
    var maxRunByClass = {};
    logData.forEach(function(r) {
      var c = Number(r[1]), run = Number(r[2]);
      if (!maxRunByClass[c] || run > maxRunByClass[c]) {
        maxRunByClass[c] = run;
      }
    });

    // 2. 이메일 또는 학생ID 매칭
    // userEmail 형식 예: s2026101@snu.ms.kr 또는 학생ID
    var targetStudent = null;
    var targetClass = null;
    var targetRun = null;

    // 이메일 앞자리(학번 등) 추출
    var emailPrefix = userEmail ? userEmail.split('@')[0].toLowerCase() : '';

    for (var i = logData.length - 1; i >= 0; i--) {
      var r = logData[i];
      var classNo = Number(r[1]);
      var run = Number(r[2]);
      var studentId = String(r[6] || '').toLowerCase();
      var studentName = String(r[5] || '');

      // 최신 회차 데이터만 검사
      if (run === maxRunByClass[classNo]) {
        if ((emailPrefix && studentId.indexOf(emailPrefix) !== -1) || studentId === userEmail.toLowerCase()) {
          targetStudent = {
            classNo: classNo,
            run: run,
            group: Number(r[3]),
            num: Number(r[4]),
            name: studentName,
            studentId: r[6],
            seat: Number(r[8]) || 0 // 0, 1, 2, 3
          };
          targetClass = classNo;
          targetRun = run;
          break;
        }
      }
    }

    if (!targetStudent) {
      return { ok: false, notFound: true, message: '학생 명단에서 일치하는 정보를 찾지 못했습니다.' };
    }

    // 3. 같은 모둠 내 짝꿍 찾기 (인접 좌석: 0번-1번 짝, 2번-3번 짝)
    var mySeat = targetStudent.seat;
    var targetPairSeat = (mySeat % 2 === 0) ? mySeat + 1 : mySeat - 1;
    var pairStudent = null;

    for (var j = 0; j < logData.length; j++) {
      var row = logData[j];
      var cNo = Number(row[1]);
      var rNo = Number(row[2]);
      var grp = Number(row[3]);
      var st = Number(row[8]) || 0;

      if (cNo === targetClass && rNo === targetRun && grp === targetStudent.group && st === targetPairSeat) {
        pairStudent = {
          num: Number(row[4]),
          name: String(row[5]),
          seat: st
        };
        break;
      }
    }

    // 고유 방 ID 생성: class_{반}_grp_{모둠}_pair_{0또는1}
    var pairGroupIndex = Math.floor(mySeat / 2);
    var roomId = 'class_' + targetClass + '_grp_' + targetStudent.group + '_pair_' + pairGroupIndex;

    return {
      ok: true,
      student: targetStudent,
      pair: pairStudent,
      roomId: roomId,
      isFirstPlayer: (mySeat % 2 === 0) // 0번, 2번 좌석이 먼저 플레이어
    };

  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/**
 * 학생들의 몬티홀 시뮬레이션 통계 결과를 시트에 영구 저장
 */
function recordMontyResults(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MONTY_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(MONTY_SHEET);
      sheet.getRange(1, 1, 1, 9)
        .setValues([['저장일시', '반', '번호', '성명', '모둠', '역할', '총시행횟수', '바꿈승률(%)', '유지승률(%)']])
        .setFontWeight('bold')
        .setBackground('#e0e7ff');
      sheet.setFrozenRows(1);
    }

    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    var row = [
      now,
      payload.classNo || '',
      payload.num || '',
      payload.name || '',
      payload.group || '',
      payload.role || '',
      payload.totalTrials || 0,
      payload.switchWinRate || 0,
      payload.stayWinRate || 0
    ];

    sheet.appendRow(row);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}
