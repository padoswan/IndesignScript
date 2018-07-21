// 선택된 페이지 여백 및 단으로 문서 여백 및 단 설정 변경 [없음] 마스터에 적용

try {
	// 현재 문서
	var swanDoc = app.activeDocument;
	// 현재 페이지(페이지 패널 더블클릭하세요)
	var swanPage = swanDoc.layoutWindows[0].activePage;
	// 현재 페이지 여백 및 단 설정
	var swanMargin = swanPage.marginPreferences;
	
	// 문서 여백 및 단 설정 변경
	swanDoc.marginPreferences.properties = {top:swanMargin.top, right:swanMargin.right, bottom:swanMargin.bottom, left:swanMargin.left, columnCount:swanMargin.columnCount, columnGutter:swanMargin.columnGutter, columnDirection:swanMargin.columnDirection, customColumns:swanMargin.customColumns, columnsPositions:swanMargin.columnsPositions};

	// 종료
	Window.alert("현재 문서의 여백 및 단 설정을 변경했습니다.\r\r위쪽: "+swanMargin.top+"  아래쪽: "+swanMargin.bottom+"\r왼쪽: "+swanMargin.left+"  오른쪽: "+swanMargin.right+"\r\열 개수: "+swanMargin.columnCount+"  열 간격: "+swanMargin.columnGutter+"\r쓰기 방향: "+colDirctionStrConvert(swanMargin.columnDirection), "알림");

	// 쓰기방향 문자변환
	function colDirctionStrConvert(dirHV) {
		var rtHVStr = ""
		if (dirHV == HorizontalOrVertical.HORIZONTAL) {
			rtHVStr = "가로";
		} else if (dirHV == HorizontalOrVertical.VERTICAL) {
			rtHVStr = "세로";
		}
		return rtHVStr;
	}
} catch (err) {
	Window.alert(err, "오류");
}
