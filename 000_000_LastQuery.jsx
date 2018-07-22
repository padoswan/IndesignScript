// 마지막으로 Grep 찾기 결과를 색인으로 등록하기

try {

	// 경고창 타이틀
	var alertTitleInfo = {en: "Infomation", ko: "알림"};
	var alertTitleWaring = {en: "Waring", ko: "경고"};
	
	// 문서 오픈 상태를 확인합니다.
	if (app.documents.length != 0) {
		// 활성문서
		var swanDoc = app.activeDocument;
		// 기본 찾기 쿼리
		main();
	} else {
		Window.alert (swanInfoMsg(1) , alertTitleInfo);
	}// 문서 오픈 상태를 확인합니다.
	
	
	//////////////////////////////////////////////////////////////////////////////////////////
	// 찾은 결과 아이템 처리
	//항상 결과가 1 이상인 경우만 이 함수가 실행함
	//////////////////////////////////////////////////////////////////////////////////////////
	function nextProgress(queryItems) {

		if (queryItems != false) {

			/////////////////////////////////////////////////////////////////////
			// 추가 함수
			/////////////////////////////////////////////////////////////////////
			
				
			// 프로그래스바 생성
			swanCreateProgressPanel(100, 400, "SwanProgress...");
			//프로그래스바 보이기
			swanProgressPanel.show();
			//프로그래스바 타이틀 변경
			swanProgressPanel.swanStaticText.text = "스크립트 실행중...";
			//기본값
			swanProgressPanel.swanProgressBar.value = 0;
			
			// 찾은 결과 반복 처리
			for (var q = 0; q < queryItems.length; q++) {
				
				//프로그래스바 수치 변경
				totalValue = ((q+1)/queryItems.length)*100;
				swanProgressPanel.swanProgressBar.value = totalValue;
				// 검색결과 아이템
				var queryItem = queryItems[q];
				
				/////////////////////////////////////////////////////////////////////
				// 추가 함수
				/////////////////////////////////////////////////////////////////////

			};//end for
		
			/////////////////////////////////////////////////////////////////////
			// 추가 함수
			/////////////////////////////////////////////////////////////////////
		
			//프로그래스바 숨기기
			swanProgressPanel.hide();
	
			// 처리 결과
			var rtMsg = "~~~~~~~~블라블라블라~~~~~~~~~~";
			return rtMsg;
		} else {
			// 처리결과 없음
			return false;
		};//end if
	};// 찾은 결과 아이템 처리
	//////////////////////////////////////////////////////////////////////////////////////////
	
	// 메인 함수
	function main() {
		// 다중 선택 유무 확인
		if (app.selection.length == 1) {
			
			var selcaseCheck = null;
			var selcaseNum = 0;
			
			// 선택 유형 확인
			switch(app.selection[0].constructor.name) {
				case "InsertionPoint":
				case "Character":
				case "Text":
				case "Word":
				case "TextStyleRange":
				case "Paragraph":
					selcaseCheck = true;
					selcaseNum = 1;
					break;
				case "TextFrame":
				case "TextColumn":
					selcaseCheck = true;
					selcaseNum = 2;
					break;
				case "Line":
				case "Cell":
				case "Table":
				case "Rectangle":
				case "Oval":
				case "Polygon":
				case "GraphicLine":
				case "Image":
				case "PDF":
				case "EPS":
					selcaseCheck = false;
					break;
				default:
					selcaseCheck = true;
					selcaseNum = 0;
					break;
			}//end switch // 선택 유형 확인

			// 선택영역 확인
			if (selcaseCheck == true) {
				if ( selcaseNum == 0) {
					swanSelect = app.documents.item(0);
				} else if ( selcaseNum == 1) {
					swanSelect = app.selection[0];
				} else if ( selcaseNum == 2) {
					swanSelect = app.selection[0].parentStory;
				}//end if
			
				// 선택 확인
				if (swanSelect.contents.length > 0) {
					// 선택한 대상에 대한 전체 진행
					selectItemProg(swanSelect);
				} else {
					Window.alert (swanInfoMsg(4, 0), alertTitleInfo);
				}//end if// 선택 확인
			
			} else {
				Window.alert (swanInfoMsg(3, swanSelect.constructor.name) , alertTitleInfo);
			}//end if// 선택영역 확인
		} else if (app.selection.length == 0) {
			// 선택이 없는경우 문서를 대상으로 함
			swanSelect = swanDoc;
			// 선택한 대상에 대한 전체 진행
			selectItemProg(swanSelect);
		} else {
			Window.alert (swanInfoMsg(2, app.selection.length) , alertTitleInfo);
		}// 다중 선택 유무 확인
	};// 메인 함수

	// 선택한 대상에 대한 전체 진행
	function selectItemProg(swanSelItem) {
		// 마지막 쿼리를 실행해서 찾은 아이템을 다음 프로세스로
		var resultProg = nextProgress(swanLastQueryGrep(swanSelItem));
						 
		// 스크립트 종료
		if (resultProg != false) {
			Window.alert (swanInfoMsg(99)+"\r\r"+resultProg, alertTitleInfo);
		} else {
			Window.alert (swanInfoMsg(98), alertTitleInfo);
		};//end if  // 스크립트 종료
	};// 선택한 대상에 대한 전체 진행

	// 찾기 함수
	function swanLastQueryGrep(rtObj) {
	
		if (app.findGrepPreferences.findWhat != ""){
			// 찾은 부분 집어넣기
			var swanFItems = rtObj.findGrep();
			
			// 찾은 결과 있는 경우
			if(swanFItems.length > 0){
				return swanFItems;
			} else {
				Window.alert (swanInfoMsg(6), alertTitleInfo);
				return false;
			}//end if
		} else {
			Window.alert (swanInfoMsg(5), alertTitleInfo);
			return false;
		}//end if
	};// 찾기 함수

	//진행바
	function swanCreateProgressPanel(swanMaximumValue, swanProgressBarWidth, swanTitleName){
		swanProgressPanel = new Window("window", swanTitleName);
		with(swanProgressPanel){
			swanProgressPanel.swanStaticText = add ("statictext", [20, 15, swanProgressBarWidth, 29], "", {scrolling:false, multiline:false});
			swanProgressPanel.swanProgressBar = add("progressbar", [12, 12, swanProgressBarWidth, 24], 0, swanMaximumValue);
		}
	};//진행바
	
	// 메시지 처리
	function swanInfoMsg(case_msg, sub_case) {
		var rt_msg = null;
		switch(case_msg) {
			case 1:// 문서 오픈 상태
				rt_msg = "현재 활성화된 인디자인 문서가 없습니다.";
				break;
			case 2:// 다중 선택 유무 확인
				if (sub_case == 0) {
					rt_msg = "스크립트 실행을 원하는 문자(텍스트프레임) 영역을 선택하세요.";
				} else {
					rt_msg = "다중 선택을 허용하지 않습니다.";
				}
				break;
			case 3:// 선택 유형 확인
				rt_msg = "현재 선택된 "+sub_case+"은(는) 문자범위 선택이 아닙니다.";
				break;
			case 4:// 선택 개체 Contents 확인
				rt_msg = "선택된 대상에는 텍스트가 없습니다.";
				break;
			case 5:// 이전쿼리
				rt_msg = "GREP으로 찾은 이전 Query가 없습니다.";
				break;
			case 6:// 검색결과 없음
				rt_msg = "이전 검색으로 찾은 결과가 없습니다.";
				break;
			case 98:// 최종종료
				rt_msg = "스크립트 실행 결과가 없습니다.";
				break;
			case 99:// 최종종료
				rt_msg = "스크립트 실행이 성공적으로 완료되었습니다.";
				break;
			default:
				rt_msg = "알수 없는 상황이 발생했습니다.";
				break;
		}
		return rt_msg;
	};// 메시지 처리

} catch (mainErr) {
		Window.alert("문제 상황:\r"+mainErr, alertTitleInfo);
}//try~catch
