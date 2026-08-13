// ============================================================
// 공공기관 보고서 AI 보조 도구 — 한글(HWP) 내장 매크로 (F10)
// ============================================================
// F9(단축키 도우미, hotkey_assistant.py)는 클립보드+전역 단축키로
// "어느 프로그램에서나" 동작하지만, 한글 안에 "진짜 버튼/메뉴"로
// 들어가 있는 느낌은 아니었다. 이 매크로는 한글 자체의 스크립트
// 매크로 기능을 이용해, 한글 메뉴/도구모음/단축키(Alt+Shift+숫자)로
// 바로 실행되고, 클립보드를 거치지 않고 문서에 직접 삽입한다.
//
// [설치 방법 — 한글에서 딱 한 번만 하면 되는 수동 작업]
// 1. 한글에서 [도구]-[매크로]-[매크로 실행] (또는 Alt+Shift+L)
// 2. 대화상자 아래 [코드 편집] 버튼 클릭 -> 스크립트 편집창이 열림
// 3. 이 파일의 내용 전체를 편집창에 붙여넣고 저장
// 4. [매크로 실행] 목록에 AI문장변환/AI회의록/AI법령검색/AI검토 가 나타나면
//    각각 선택 후 원하는 단축키(Alt+Shift+1~4 등)를 지정
// 5. (선택) [도구]-[사용자 설정]-[메뉴 및 도구 상자] 에서 "매크로" 항목을
//    찾아 도구모음에 아이콘으로 추가하면 진짜 버튼처럼 쓸 수 있다.
//    (4~5단계의 정확한 화면 구성은 한글 버전마다 다를 수 있어, 이 부분은
//    직접 한글 안에서 확인이 필요하다 — Claude는 데스크톱 앱의 화면을
//    직접 클릭할 수 없어 이 문서로만 안내한다.)
//
// [기술적으로 확인된 부분 — Python win32com으로 사전 검증됨]
// - GetTextFile("TEXT", "saveblock") : 현재 선택 영역의 텍스트를 그대로 가져옴
// - HAction / HParameterSet.HInsertText : 문서에 텍스트를 직접 삽입 (클립보드 불필요)
// - 매크로 스크립트 안에서도 HAction/HParameterSet/GetTextFile 은 앞에
//   아무것도 안 붙이고(hwp. 같은 접두사 없이) 바로 쓸 수 있는 전역 객체다.
// - ActiveXObject("WScript.Shell")로 외부 프로그램(파이썬) 호출 가능
//   (단, 한글이 처음 실행할 때 "매크로 보안 경고" 창을 띄울 수 있음 — 정상 동작)
// - 한글<->파이썬 간 텍스트 교환은 UTF-16(유니코드) 임시 파일로 한다.
//   (클립보드/커맨드라인 인자보다 한글 인코딩이 깨질 위험이 없음. 실제로
//    "TEXT"/커맨드라인 방식은 한글이 깨지는 걸 확인해서 이 방식으로 결정)
// ============================================================

// ↓↓↓ 사용자 PC 환경에 맞게 필요시 수정 ↓↓↓
var PY_EXE = "C:\\Users\\lakka\\AppData\\Local\\Programs\\Python\\Python313\\python.exe";
var PROJECT_DIR = "D:\\보고서자동화";
var BRIDGE = PROJECT_DIR + "\\macro_bridge.py";
// ↑↑↑ 사용자 PC 환경에 맞게 필요시 수정 ↑↑↑


function OnScriptMacro_AI문장변환()
{
	RunAIBridge("gaejoshik", true);
}

function OnScriptMacro_AI회의록()
{
	RunAIBridge("meeting", true);
}

function OnScriptMacro_AI법령검색()
{
	RunAIBridge("law", true);
}

function OnScriptMacro_AI검토()
{
	RunAIBridge("review", false);
}


// mode: "gaejoshik" | "meeting" | "law" | "review" (macro_bridge.py 의 MODES 와 일치해야 함)
// insertBelow: true 면 결과를 원문 아래 줄에 삽입, false 면 알림창으로만 표시 (문서는 안 바뀜)
function RunAIBridge(mode, insertBelow)
{
	var selected = GetTextFile("TEXT", "saveblock");
	if (!selected || selected.length == 0)
	{
		ShowPopup("먼저 문장을 드래그로 선택하세요.", "AI 도우미", 0);
		return;
	}

	// 처리 시작 알림 (3초 후 자동으로 사라짐 - 클릭 안 해도 됨)
	// "멈춘 건가 처리 중인가" 헷갈리지 않도록 미리 알려주는 용도.
	ShowPopup("AI 처리를 시작합니다.\n15~30초 정도 걸릴 수 있어요.\n(한글 화면이 잠깐 멈춘 것처럼 보여도 정상입니다)", "처리 중...", 3);

	var shell = new ActiveXObject("WScript.Shell");
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var tempDir = shell.ExpandEnvironmentStrings("%TEMP%");
	var inPath = tempDir + "\\hwp_ai_in.txt";
	var outPath = tempDir + "\\hwp_ai_out.txt";

	if (fso.FileExists(outPath))
		fso.DeleteFile(outPath);

	// true, true = 덮어쓰기, 유니코드(UTF-16)로 저장
	var tsIn = fso.CreateTextFile(inPath, true, true);
	tsIn.Write(selected);
	tsIn.Close();

	var cmd = '"' + PY_EXE + '" "' + BRIDGE + '" ' + mode + ' "' + inPath + '" "' + outPath + '"';
	// 0 = 창 숨김, true = 끝날 때까지 대기 (LLM 응답까지 보통 15~30초 걸림)
	shell.Run(cmd, 0, true);

	if (!fso.FileExists(outPath))
	{
		ShowPopup("AI 처리에 실패했습니다.\n(파이썬 경로, Ollama 실행 여부를 확인하세요)", "AI 도우미 오류", 0);
		return;
	}

	// 1 = ForReading, -1 = TristateTrue(유니코드로 열기)
	var tsOut = fso.OpenTextFile(outPath, 1, false, -1);
	var result = tsOut.ReadAll();
	tsOut.Close();

	if (insertBelow)
	{
		HAction.Run("MoveLineEnd");   // 선택을 해제하고 그 줄 끝으로 이동
		HAction.Run("BreakPara");     // 줄바꿈
		HParameterSet.HInsertText.Text = result;
		HAction.Execute("InsertText", HParameterSet.HInsertText.HSet);
		ShowPopup("완료되었습니다.", "AI 도우미", 2);   // 2초 후 자동으로 사라짐
	}
	else
	{
		ShowPopup(result, "AI 검토 결과 (문서는 변경되지 않음)", 0);
	}
}

// secondsToWait: 0이면 사용자가 닫을 때까지 유지, 그 외엔 그 초 수만큼 있다가 자동으로 닫힘
function ShowPopup(message, title, secondsToWait)
{
	var shell = new ActiveXObject("WScript.Shell");
	shell.Popup(message, secondsToWait, title, 64);
}
