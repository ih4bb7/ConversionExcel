function selectChange() {
    var activeElement = document.activeElement;
    var value = activeElement.value;
    var processCount = activeElement.id.substring(6);
    var id1 = "argument1_" + processCount;
    var id2 = "argument2_" + processCount;
    var id3 = "argument3_" + processCount;
    var id4 = "argument4_" + processCount;
    var id5 = "argument5_" + processCount;
    switch (value) {
        case "":
            readOnly5(id1, id2, id3, id4, id5);
            break;
        case "書き込み":
            readOnly2(id1, id2, id3, id4, id5);
            placeholder3(id1, id2, id3, "シート名", "セル番地", "値");
            break;
        case "セルコピペ":
            readOnly1(id1, id2, id3, id4, id5);
            placeholder4(id1, id2, id3, id4, "読み込みシート名", "セル番地", "書き込みシート名", "セル番地");
            break;
        case "行コピペ":
            readOnly1(id1, id2, id3, id4, id5);
            placeholder4(id1, id2, id3, id4, "読み込みシート名", "行番号", "書き込みシート名", "行番号");
            break;
        case "数字書き込み":
            readOnly2(id1, id2, id3, id4, id5);
            placeholder3(id1, id2, id3, "シート名", "セル番地", "数字");
            break;
        case "関数書き込み":
            readOnly2(id1, id2, id3, id4, id5);
            placeholder3(id1, id2, id3, "シート名", "セル番地", "関数");
            break;
        default:
            break;
    }
}

function placeholder0(id1, id2, id3, id4, id5) {
    document.getElementById(id1).placeholder = "";
    document.getElementById(id2).placeholder = "";
    document.getElementById(id3).placeholder = "";
    document.getElementById(id4).placeholder = "";
    document.getElementById(id5).placeholder = "";
}

function placeholder1(id1, placeholder1) {
    document.getElementById(id1).placeholder = placeholder1;
}

function placeholder2(id1, id2, placeholder1, placeholder2) {
    document.getElementById(id1).placeholder = placeholder1;
    document.getElementById(id2).placeholder = placeholder2;
}

function placeholder3(id1, id2, id3, placeholder1, placeholder2, placeholder3) {
    document.getElementById(id1).placeholder = placeholder1;
    document.getElementById(id2).placeholder = placeholder2;
    document.getElementById(id3).placeholder = placeholder3;
}

function placeholder4(id1, id2, id3, id4, placeholder1, placeholder2, placeholder3, placeholder4) {
    document.getElementById(id1).placeholder = placeholder1;
    document.getElementById(id2).placeholder = placeholder2;
    document.getElementById(id3).placeholder = placeholder3;
    document.getElementById(id4).placeholder = placeholder4;
}

function placeholder5(id1, id2, id3, id4, id5, placeholder1, placeholder2, placeholder3, placeholder4, placeholder5) {
    document.getElementById(id1).placeholder = placeholder1;
    document.getElementById(id2).placeholder = placeholder2;
    document.getElementById(id3).placeholder = placeholder3;
    document.getElementById(id4).placeholder = placeholder4;
    document.getElementById(id5).placeholder = placeholder5;
}

function readOnly0(id1, id2, id3, id4, id5) {
    document.getElementById(id1).value = "";
    document.getElementById(id2).value = "";
    document.getElementById(id3).value = "";
    document.getElementById(id4).value = "";
    document.getElementById(id5).value = "";
    document.getElementById(id1).readOnly = false;
    document.getElementById(id2).readOnly = false;
    document.getElementById(id3).readOnly = false;
    document.getElementById(id4).readOnly = false;
    document.getElementById(id5).readOnly = false;
    placeholder5(id1, id2, id3, id4, id5, "", "", "", "", "");
}

function readOnly1(id1, id2, id3, id4, id5) {
    document.getElementById(id1).value = "";
    document.getElementById(id2).value = "";
    document.getElementById(id3).value = "";
    document.getElementById(id4).value = "";
    document.getElementById(id5).value = "";
    document.getElementById(id1).readOnly = false;
    document.getElementById(id2).readOnly = false;
    document.getElementById(id3).readOnly = false;
    document.getElementById(id4).readOnly = false;
    document.getElementById(id5).readOnly = true;
    placeholder5(id1, id2, id3, id4, id5, "", "", "", "", "");
}

function readOnly2(id1, id2, id3, id4, id5) {
    document.getElementById(id1).value = "";
    document.getElementById(id2).value = "";
    document.getElementById(id3).value = "";
    document.getElementById(id4).value = "";
    document.getElementById(id5).value = "";
    document.getElementById(id1).readOnly = false;
    document.getElementById(id2).readOnly = false;
    document.getElementById(id3).readOnly = false;
    document.getElementById(id4).readOnly = true;
    document.getElementById(id5).readOnly = true;
    placeholder5(id1, id2, id3, id4, id5, "", "", "", "", "");
}

function readOnly3(id1, id2, id3, id4, id5) {
    document.getElementById(id1).value = "";
    document.getElementById(id2).value = "";
    document.getElementById(id3).value = "";
    document.getElementById(id4).value = "";
    document.getElementById(id5).value = "";
    document.getElementById(id1).readOnly = false;
    document.getElementById(id2).readOnly = false;
    document.getElementById(id3).readOnly = true;
    document.getElementById(id4).readOnly = true;
    document.getElementById(id5).readOnly = true;
    placeholder5(id1, id2, id3, id4, id5, "", "", "", "", "");
}

function readOnly4(id1, id2, id3, id4, id5) {
    document.getElementById(id1).value = "";
    document.getElementById(id2).value = "";
    document.getElementById(id3).value = "";
    document.getElementById(id4).value = "";
    document.getElementById(id5).value = "";
    document.getElementById(id1).readOnly = false;
    document.getElementById(id2).readOnly = true;
    document.getElementById(id3).readOnly = true;
    document.getElementById(id4).readOnly = true;
    document.getElementById(id5).readOnly = true;
    placeholder5(id1, id2, id3, id4, id5, "", "", "", "", "");
}

function readOnly5(id1, id2, id3, id4, id5) {
    document.getElementById(id1).value = "";
    document.getElementById(id2).value = "";
    document.getElementById(id3).value = "";
    document.getElementById(id4).value = "";
    document.getElementById(id5).value = "";
    document.getElementById(id1).readOnly = true;
    document.getElementById(id2).readOnly = true;
    document.getElementById(id3).readOnly = true;
    document.getElementById(id4).readOnly = true;
    document.getElementById(id5).readOnly = true;
    placeholder5(id1, id2, id3, id4, id5, "", "", "", "", "");
}