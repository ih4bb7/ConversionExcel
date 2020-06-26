﻿function selectChange() {
    var activeElement = document.activeElement;
    var value = activeElement.value;
    var processCount = activeElement.id.substring(6);
    var id1 = "argument-" + processCount + "-1";
    var id2 = "argument-" + processCount + "-2";
    var id3 = "argument-" + processCount + "-3";
    var id4 = "argument-" + processCount + "-4";
    var id5 = "argument-" + processCount + "-5";
    if (value == "") {
        readOnly5(id1, id2, id3, id4, id5);
    }
    else if (value == "書き込み") {
        readOnly2(id1, id2, id3, id4, id5);
        placeholder3(id1, id2, id3, "シート名", "セル番地", "値");
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
    document.getElementById(id1).readOnly = false;
    document.getElementById(id2).readOnly = false;
    document.getElementById(id3).readOnly = false;
    document.getElementById(id4).readOnly = false;
    document.getElementById(id5).readOnly = false;
    placeholder5(id1, id2, id3, id4, id5, "", "", "", "", "");
}

function readOnly1(id1, id2, id3, id4, id5) {
    document.getElementById(id1).readOnly = false;
    document.getElementById(id2).readOnly = false;
    document.getElementById(id3).readOnly = false;
    document.getElementById(id4).readOnly = false;
    document.getElementById(id5).readOnly = true;
    placeholder4(id1, id2, id3, id4, "", "", "", "");
}

function readOnly2(id1, id2, id3, id4, id5) {
    document.getElementById(id1).readOnly = false;
    document.getElementById(id2).readOnly = false;
    document.getElementById(id3).readOnly = false;
    document.getElementById(id4).readOnly = true;
    document.getElementById(id5).readOnly = true;
    placeholder3(id1, id2, id3, "", "", "");
}

function readOnly3(id1, id2, id3, id4, id5) {
    document.getElementById(id1).readOnly = false;
    document.getElementById(id2).readOnly = false;
    document.getElementById(id3).readOnly = true;
    document.getElementById(id4).readOnly = true;
    document.getElementById(id5).readOnly = true;
    placeholder2(id1, id2, "", "");
}

function readOnly4(id1, id2, id3, id4, id5) {
    document.getElementById(id1).readOnly = false;
    document.getElementById(id2).readOnly = true;
    document.getElementById(id3).readOnly = true;
    document.getElementById(id4).readOnly = true;
    document.getElementById(id5).readOnly = true;
    placeholder1(id1, "");
}

function readOnly5(id1, id2, id3, id4, id5) {
    document.getElementById(id1).readOnly = true;
    document.getElementById(id2).readOnly = true;
    document.getElementById(id3).readOnly = true;
    document.getElementById(id4).readOnly = true;
    document.getElementById(id5).readOnly = true;
    placeholder0(id1, id2, id3, id4, id5);
}