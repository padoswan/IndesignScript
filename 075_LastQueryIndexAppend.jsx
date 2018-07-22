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
			
			// 등록된 색인 카운트
			var sucessIndexCount = 0;
			// 색인 정렬 그룹
			var indexGroupArr = new Array();
			// 기호는 항상 표시되며, 이 속성에서 true로 변경되면 해당 색인 정렬 표시를 하지만, 
			// 표시된 정렬그룹을 turn off 하지는 않는다.
			// 중국어는 병음/획수 처리하지 않고 기호로 등록함.
			
			//  첫글자가 영어/한글/한자/일본어/키릴자모/그리스어 경우 정렬순서에 불필요한 기호문자들 지우기
			var rtSortTextTypeArr = swanIndexSortTextDlg();
			
			//선택 대화창에서 실행 선택한 경우
			if (rtSortTextTypeArr[0] == true) {
				
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
					// 색인등록, 정렬기준에 기호를 삭제한 텍스트 사용 여부
					resultIndex = swanIndexAppend(queryItem, rtSortTextTypeArr[1]);
					if (resultIndex[0] == true) {
						// 등록된 색인 카운트
						sucessIndexCount++;
						if (indexGroupArr.toString().indexOf(resultIndex[1]) == -1) {
							indexGroupArr.push(resultIndex[1]);
						};//end if
					};//end if
				};//end for
			
				/////////////////////////////////////////////////////////////////////
				// 추가 함수
				/////////////////////////////////////////////////////////////////////
				
				// 색인 정렬 표시 추가
				if (indexGroupArr.length > 0) {
					for (var gr = 0; gr < indexGroupArr.length; gr++) {
						var indexGroupArrStr =indexGroupArr[gr];
						if (indexGroupArrStr == "kIndexGroup_Numeric") {
							swanDoc.indexingSortOptions.item(app.translateKeyString("$ID/kIndexGroup_Numeric")).include = true;
						} else if (indexGroupArrStr =="kIndexGroup_Alphabet") {
							swanDoc.indexingSortOptions.item(app.translateKeyString("$ID/kIndexGroup_Alphabet")).include = true;
						} else if (indexGroupArrStr == "kIndexGroup_Korean") {
							swanDoc.indexingSortOptions.item(app.translateKeyString("$ID/kIndexGroup_Korean")).include = true;
						} else if (indexGroupArrStr == "kIndexGroup_Kana") {
							swanDoc.indexingSortOptions.item(app.translateKeyString("$ID/kIndexGroup_Kana")).include = true;
						} else if (indexGroupArrStr == "kIndexGroup_Chinese") {
							swanDoc.indexingSortOptions.item(app.translateKeyString("$ID/kIndexGroup_Chinese")).include = true;
						} else if (indexGroupArrStr == "kWRIndexGroup_CyrillicAlphabet") {
							swanDoc.indexingSortOptions.item(app.translateKeyString("$ID/kWRIndexGroup_CyrillicAlphabet")).include = true;
						} else if (indexGroupArrStr == "kWRIndexGroup_GreekAlphabet") {
							swanDoc.indexingSortOptions.item(app.translateKeyString("$ID/kWRIndexGroup_GreekAlphabet")).include = true;
						} else if (indexGroupArrStr == "kWRIndexGroup_ArabicAlphabet") {
							swanDoc.indexingSortOptions.item(app.translateKeyString("$ID/kWRIndexGroup_ArabicAlphabet")).include = true;
						} else if (indexGroupArrStr == "kWRIndexGroup_HebrewAlphabet") {
							swanDoc.indexingSortOptions.item(app.translateKeyString("$ID/kWRIndexGroup_HebrewAlphabet")).include = true;
						};// end if
					};// end for
				};// end if // 색인 정렬 표시 추가
			
				//프로그래스바 숨기기
				swanProgressPanel.hide();
		
				// 처리 결과
				var rtMsg = "색인 등록 대상 '"+queryItems.length+"'개 중에 모두 '"+sucessIndexCount+"'개가 등록되었습니다.";
				return rtMsg;
			} else {
				// 처리결과 없음
				return false;
			};//end if
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

	// 색인등록
	function swanIndexAppend(queryItemObj, clearTextTF) {
		
		// CJK한자
		swanCJK_Hanja = {"F900":"8C48", "F901":"66F4", "F902":"8ECA", "F903":"8CC8", "F904":"6ED1", "F905":"4E32", "F906":"53E5", "F907":"9F9C", "F908":"9F9C", "F909":"5951", "F90A":"91D1", "F90B":"5587", "F90C":"5948", "F90D":"61F6", "F90E":"7669", "F90F":"7F85", "F910":"863F", "F911":"87BA", "F912":"88F8", "F913":"908F", "F914":"6A02", "F915":"6D1B", "F916":"70D9", "F917":"73DE", "F918":"843D", "F919":"916A", "F91A":"99F1", "F91B":"4E82", "F91C":"5375", "F91D":"6B04", "F91E":"721B", "F91F":"862D", "F920":"9E1E", "F921":"5D50", "F922":"6FEB", "F923":"85CD", "F924":"8964", "F925":"62C9", "F926":"81D8", "F927":"881F", "F928":"5ECA", "F929":"6717", "F92A":"6D6A", "F92B":"72FC", "F92C":"90DE", "F92D":"4F86", "F92E":"51B7", "F92F":"52DE", "F930":"64C4", "F931":"6AD3", "F932":"7210", "F933":"76E7", "F934":"8001", "F935":"8606", "F936":"865C", "F937":"8DEF", "F938":"9732", "F939":"9B6F", "F93A":"9DFA", "F93B":"788C", "F93C":"797F", "F93D":"7DA0", "F93E":"83C9", "F93F":"9304", "F940":"9E7F", "F941":"8AD6", "F942":"58DF", "F943":"5F04", "F944":"7C60", "F945":"807E", "F946":"7262", "F947":"78CA", "F948":"8CC2", "F949":"96F7", "F94A":"58D8", "F94B":"5C62", "F94C":"6A13", "F94D":"6DDA", "F94E":"6F0F", "F94F":"7D2F", "F950":"7E37", "F951":"964B", "F952":"52D2", "F953":"808B", "F954":"51DC", "F955":"51CC", "F956":"7A1C", "F957":"7DBE", "F958":"83F1", "F959":"9675", "F95A":"8B80", "F95B":"62CF", "F95C":"6A02", "F95D":"8AFE", "F95E":"4E39", "F95F":"5BE7", "F960":"6012", "F961":"7387", "F962":"7570", "F963":"5317", "F964":"78FB", "F965":"4FBF", "F966":"5FA9", "F967":"4E0D", "F968":"6CCC", "F969":"6578", "F96A":"7D22", "F96B":"53C3", "F96C":"585E", "F96D":"7701", "F96E":"8449", "F96F":"8AAA", "F970":"6BBA", "F971":"8FB0", "F972":"6C88", "F973":"62FE", "F974":"82E5", "F975":"63A0", "F976":"7565", "F977":"4EAE", "F978":"5169", "F979":"51C9", "F97A":"6881", "F97B":"7CE7", "F97C":"826F", "F97D":"8AD2", "F97E":"91CF", "F97F":"52F5", "F980":"5442", "F981":"5973", "F982":"5EEC", "F983":"65C5", "F984":"6FFE", "F985":"792A", "F986":"95AD", "F987":"9A6A", "F988":"9E97", "F989":"9ECE", "F98A":"529B", "F98B":"66C6", "F98C":"6B77", "F98D":"8F62", "F98E":"5E74", "F98F":"6190", "F990":"6200", "F991":"649A", "F992":"6F23", "F993":"7149", "F994":"7489", "F995":"79CA", "F996":"7DF4", "F997":"806F", "F998":"8F26", "F999":"84EE", "F99A":"9023", "F99B":"934A", "F99C":"5217", "F99D":"52A3", "F99E":"54BD", "F99F":"70C8", "F9A0":"88C2", "F9A1":"8AAA", "F9A2":"5EC9", "F9A3":"5FF5", "F9A4":"637B", "F9A5":"6BAE", "F9A6":"7C3E", "F9A7":"7375", "F9A8":"4EE4", "F9A9":"56F9", "F9AA":"5BE7", "F9AB":"5DBA", "F9AC":"601C", "F9AD":"73B2", "F9AE":"7469", "F9AF":"7F9A", "F9B0":"8046", "F9B1":"9234", "F9B2":"96F6", "F9B3":"9748", "F9B4":"9818", "F9B5":"4F8B", "F9B6":"79AE", "F9B7":"91B4", "F9B8":"96B7", "F9B9":"60E1", "F9BA":"4E86", "F9BB":"50DA", "F9BC":"5BEE", "F9BD":"5C3F", "F9BE":"6599", "F9BF":"6A02", "F9C0":"71CE", "F9C1":"7642", "F9C2":"84FC", "F9C3":"907C", "F9C4":"9F8D", "F9C5":"6688", "F9C6":"962E", "F9C7":"5289", "F9C8":"677B", "F9C9":"67F3", "F9CA":"6D41", "F9CB":"6E9C", "F9CC":"7409", "F9CD":"7559", "F9CE":"786B", "F9CF":"7D10", "F9D0":"985E", "F9D1":"516D", "F9D2":"622E", "F9D3":"9678", "F9D4":"502B", "F9D5":"5D19", "F9D6":"6DEA", "F9D7":"8F2A", "F9D8":"5F8B", "F9D9":"6144", "F9DA":"6817", "F9DB":"7387", "F9DC":"9686", "F9DD":"5229", "F9DE":"540F", "F9DF":"5C65", "F9E0":"6613", "F9E1":"674E", "F9E2":"68A8", "F9E3":"6CE5", "F9E4":"7406", "F9E5":"75E2", "F9E6":"7F79", "F9E7":"88CF", "F9E8":"88E1", "F9E9":"91CC", "F9EA":"96E2", "F9EB":"533F", "F9EC":"6EBA", "F9ED":"541D", "F9EE":"71D0", "F9EF":"7498", "F9F0":"85FA", "F9F1":"96A3", "F9F2":"9C57", "F9F3":"9E9F", "F9F4":"6797", "F9F5":"6DCB", "F9F6":"81E8", "F9F7":"7ACB", "F9F8":"7B20", "F9F9":"7C92", "F9FA":"72C0", "F9FB":"7099", "F9FC":"8B58", "F9FD":"4EC0", "F9FE":"8336", "F9FF":"523A", "FA00":"5207", "FA01":"5EA6", "FA02":"62D3", "FA03":"7CD6", "FA04":"5B85", "FA05":"6D1E", "FA06":"66B4", "FA07":"8F3B", "FA08":"884C", "FA09":"964D", "FA0A":"898B", "FA0B":"5ED3", "FA0C":"5140", "FA0D":"55C0", "FA0E":"", "FA0F":"5502", "FA10":"585A", "FA11":"", "FA12":"6674", "FA13":"", "FA14":"", "FA15":"51DE", "FA16":"732A", "FA17":"76CA", "FA18":"793C", "FA19":"795E", "FA1A":"7965", "FA1B":"798F", "FA1C":"9756", "FA1D":"7CBE", "FA1E":"7FBD", "FA1F":"", "FA20":"", "FA21":"", "FA22":"8AF8", "FA23":"", "FA24":"", "FA25":"9038", "FA26":"90FD", "FA27":"", "FA28":"", "FA29":"", "FA2A":"98EF", "FA2B":"98FC", "FA2C":"9928", "FA2D":"9DB4", "FA2E":"", "FA2F":"", "FA30":"4FAE", "FA31":"50E7", "FA32":"514D", "FA33":"52C9", "FA34":"52E4", "FA35":"5351", "FA36":"559D", "FA37":"5606", "FA38":"5668", "FA39":"5840", "FA3A":"58A8", "FA3B":"5C64", "FA3C":"5C6E", "FA3D":"6094", "FA3E":"", "FA3F":"618E", "FA40":"61F2", "FA41":"654F", "FA42":"65E2", "FA43":"6691", "FA44":"6885", "FA45":"6D77", "FA46":"6E1A", "FA47":"6F22", "FA48":"716E", "FA49":"722B", "FA4A":"7422", "FA4B":"7891", "FA4C":"793E", "FA4D":"7949", "FA4E":"7948", "FA4F":"7950", "FA50":"7956", "FA51":"795D", "FA52":"798D", "FA53":"798E", "FA54":"7A40", "FA55":"7A81", "FA56":"7BC0", "FA57":"7DF4", "FA58":"7E09", "FA59":"7E41", "FA5A":"7F72", "FA5B":"8005", "FA5C":"81ED", "FA5D":"8279", "FA5E":"8279", "FA5F":"8457", "FA60":"8910", "FA61":"8996", "FA62":"8B01", "FA63":"8B39", "FA64":"8CD3", "FA65":"8D08", "FA66":"8FB6", "FA67":"9038", "FA68":"96E3", "FA69":"97FF", "FA6A":"983B"};

		// 기본색인유무
		if (swanDoc.indexes.length == 0) {
			swanDoc.indexes.add();
		}// 기본색인유무

		// 기본색인
		if (swanDoc.indexes.length == 1) {
			swanIndex = swanDoc.indexes[0];
		};
		
		// 색인 등록 표제어
		var sourTopicName = queryItemObj.contents;
		// 색인 정렬 표지어
		var clearTopicName = "";
		
		// 정렬용 문자 처리
		for (var tp = 0; tp < sourTopicName.length; tp++) {
			var swanChInt = sourTopicName[tp].charCodeAt(0);
			var swanChHexa = swanHexString(sourTopicName[tp]);
			
			// 영역 변환
			if (0xFF00 <= swanChInt && swanChInt <= 0xFF5E) {
				// 전각=>반각변환
				clearTopicName += String.fromCharCode(swanChInt-0xFEE0);
			} else if (0xF900 <= swanChInt && swanChInt <= 0xFA6A) {
				//CJK호환용한자변환
				clearTopicName += swanCharString(swanCJK_Hanja[swanChHexa]);
			} else {
				clearTopicName += sourTopicName[tp];
			};//end if
		};//end for // 정렬용 문자 처리
	
		// 앞뒤 공백 및 내부 공백지우기
		clearTopicName = clearTopicName.replace(/\s+/gi, "");
		
		// 첫글자에 따른 색인 정렬 유무 체크
		// 한자는 기호로 처리함 
		// 중국어 색인 처리는 지원하지 않음
		
		// 색인 정렬용 첫문자 영역 확인
		var clearFirstCharCodeInt = clearTopicName[0].charCodeAt(0);
		
		// 불필요한 기호문자 지움 설정
		var defSignCharDel = true;
		// 색인 정렬 옵션
		var indexSortKeyString = "";
		
		// 첫글자가 영어/한글/한자/일본어/키릴자모/그리스어 경우 정렬순서에 불필요한 기호문자들 지우기
		if (0x0030 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x0039) {
			// 숫자
			indexSortKeyString = "kIndexGroup_Numeric";
		} else if (0x0041 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x005A) {
			// 영어 대문자
			indexSortKeyString = "kIndexGroup_Alphabet";
		} else if (0x0061 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x007A) {
			// 영어 소문자
			indexSortKeyString = "kIndexGroup_Alphabet";
		} else if (0x00C0 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x00D6) {
			// ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ
			indexSortKeyString = "kIndexGroup_Alphabet";
		} else if (0x00D9 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x00F6) {
			// ÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö
			indexSortKeyString = "kIndexGroup_Alphabet";
		} else if (0x00F9 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x00FF) {
			// ùúûüýþÿ
			indexSortKeyString = "kIndexGroup_Alphabet";
		} else if (0x0100 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x017F) {
			// 라틴확장A
			indexSortKeyString = "kIndexGroup_Alphabet";
		} else if (0x0370 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x03FF) {
			// 그리스어와 콥트어
			indexSortKeyString = "kWRIndexGroup_GreekAlphabet";
		} else if (0x0400 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x052F) {
			// 키릴자모
			indexSortKeyString = "kWRIndexGroup_CyrillicAlphabet";
		} else if (0x0590 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x05EF) {
			// 히브리어
			indexSortKeyString = "kWRIndexGroup_HebrewAlphabet";
		} else if (0x0600 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x06FF) {
			// 아랍어
			indexSortKeyString = "kWRIndexGroup_ArabicAlphabet";
		} else if (0x0750 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x077F) {
			// 아랍어 보충 => 색인 기호로 등록됨
			indexSortKeyString = "kIndexGroup_Symbol";
		} else if (0x1100 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x11FF) {
			// 한글 자모 => 색인 기호로 등록됨
			indexSortKeyString = "kIndexGroup_Symbol";
		} else if (0x0180 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x024F) {
			// 라틴어B => 색인 기호로 등록됨
			indexSortKeyString = "kIndexGroup_Symbol";
		} else if (0x1E00 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x1EFF) {
			// 라틴어추가확장 => 색인 기호로 등록됨
			indexSortKeyString = "kIndexGroup_Symbol";
		} else if (0x1F00 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x1FFF) {
			// 그리스어 확장 => 색인 기호로 등록됨
			indexSortKeyString = "kIndexGroup_Symbol";
		} else if (0x2C60 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x2C7F) {
			// 라틴어 C => 색인 기호로 등록됨
			indexSortKeyString = "kIndexGroup_Symbol";
		} else if (0x2E80 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x2EFF) {
			// 한중일 부수 보충 => 색인 별도 처리가 필요한 영역 => 색인 기호로 등록됨
			indexSortKeyString = "kIndexGroup_Symbol";
		} else if (0x2F00 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x2FDF) {
			// 강희자전 부수 => 색인 별도 처리가 필요한 영역 => 색인 기호로 등록됨
			indexSortKeyString = "kIndexGroup_Symbol";
		} else if (0x3040 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x309F) {
			// 히라가나
			indexSortKeyString = "kIndexGroup_Kana";
		} else if (0x30A0 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x30FF) {
			// 가타가나
			indexSortKeyString = "kIndexGroup_Kana";
		} else if (0x3130 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x318F) {
			// 호환용 한글 자모
			indexSortKeyString = "kIndexGroup_Korean";
		} else if (0x3400 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x4DBF) {
			// 한중일 통합한자 A => 색인 별도 처리가 필요한 영역 => 색인 기호로 등록됨
			indexSortKeyString = "kIndexGroup_Symbol";
		} else if (0x4E00 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0x9FFF) {
			// 한중일 통합한자 => 색인 별도 처리가 필요한 영역 => 색인 기호로 등록됨
			indexSortKeyString = "kIndexGroup_Symbol";
		} else if (0xAC00 <= clearFirstCharCodeInt && clearFirstCharCodeInt <= 0xD7AF) {
			// 한글 자모
			indexSortKeyString = "kIndexGroup_Korean";
		} else {
			// 설정영역이 아닌 경우 기호문자 지우지 않음
			defSignCharDel = false;
			indexSortKeyString = "kIndexGroup_Symbol";
		}//end if // 첫글자가 영어/한글/한자/일본어/키릴자모/그리스어 경우 정렬순서에 불필요한 기호문자들 지우기

		// // 불필요한 기호문자 지움 설정
		if (defSignCharDel == true && clearTextTF == true) {
			clearTopicName = clearTopicName.replace(/[\u0021-\u002F\u003A-\u0040\u005B-\u0060\u007B-\u007F]/gi, "");
		}//end if

		// 색인 등록에 오류 허용
		try {
			// 색인 항목 등록
			insertTopic = swanIndex.topics.add(sourTopicName, clearTopicName);
			// 페이지참조
			insertTopic.pageReferences.add (queryItemObj,PageReferenceType.CURRENT_PAGE);
			
			return [true, indexSortKeyString];
		} catch (addErr) {
			return [false, false];
		}; // 색인 등록에 오류 허용
	};// 색인 등록

	// 색인 처리 부가
	// Hex->Char
	function swanCharString(swanCtHexCode) {
		tmp_charInt = parseInt(swanCtHexCode, 16);
		tmp_string = String.fromCharCode(tmp_charInt);
		return tmp_string;
	};// Hex->Char

	// 색인 처리 부가
	// Char->Hex
	function swanHexString(swanCtCharCode) {
		tmp_charInt = swanCtCharCode.charCodeAt(0);
		tmp_charHex = tmp_charInt.toString(16);
		tmp_string = tmp_charHex.toUpperCase();
		return tmp_string;
	};// Char->Hex

	// 선택 DLG
	function swanIndexSortTextDlg(){
		var swDialog = new Window( "dialog", "정렬 텍스트 사용자 설정" ); 

		// 선택
		var rButtonPanel = swDialog.add("panel", [0, 0, 400, 120], "정렬 텍스트 처리 선택");
		rButtonPanel.sourceRdButton = rButtonPanel.add("radiobutton", [30, 30, 380, 45], "검색 텍스트를 정렬 기준으로 사용");
		rButtonPanel.clearRdButton = rButtonPanel.add("radiobutton", [30, 65, 380, 80], "기호가 삭제된 텍스트를 정렬기준으로 사용");
	
		// 변환과 취소 버튼 그룹
		var gButton = swDialog.add("group",  [0, 140, 400, 195]); 
		swDialog.cvtButton = gButton.add("button", [190, 0, 290, 45], "실행", {name: "execute"});
		swDialog.canButton = gButton.add("button", [300, 0, 400, 45], "취소", {name:"cancel"}); 

		//기본 설정
		rButtonPanel.clearRdButton.value = true;
		var clearType = true;
	
		// 변환과 취소 버튼 그룹 선택에 따른 실행 분기
		swDialog.cvtButton.onClick = function() { 
			if (rButtonPanel.clearRdButton.value == true) {
				clearType = true;
			} else {
				clearType = false;
			}
			this.window.close(true); 
		} 	
		swDialog.canButton.onClick = function() { 
			this.window.close(false); 
		}
	
		// 대화창 위치
		swDialog.center();
		
		//DLG BOX 표시
		var swDlgResult = swDialog.show();
		
		if (swDlgResult == true) {
			return [true,  clearType]; 
			swDialog.destroy();
		} else {
			return [false];
			swDialog.destroy();
		}
	};// 선택 DLG
	
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
