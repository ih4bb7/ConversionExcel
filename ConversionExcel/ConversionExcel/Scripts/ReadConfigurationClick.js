$(function () {
    $("#readConfiguration").click(function () {
        $.ajax({
            data: { path: $('#configurationPath').val() },
            type: "POST",
            url: "/Home/readConfiguration_Click",
            success: function (result) {
                if (!result.result.IsFile || result.result.HasError) {
                    alert(result.result.Message);
                    return false;
                }
                var processCount = (document.getElementById('Processes').innerHTML.match(/処理内容/g) || []).length;
                for (var i = 1; i < processCount + 1; i++) {
                    if (i == 1) continue;
                    $('#process_' + i).remove();
                }
                document.getElementById('readPath').value = result.result.Parent.ReadPath;
                document.getElementById('outputPath').value = result.result.Parent.OutputPath;
                for (var i = 1; i < result.result.Parent.Processes.length + 1; i++) {
                    document.getElementById('shori_' + i).value = result.result.Parent.Processes[i - 1].Shori;
                    var id1 = 'argument1_' + i;
                    var id2 = 'argument2_' + i;
                    var id3 = 'argument3_' + i;
                    var id4 = 'argument4_' + i;
                    var id5 = 'argument5_' + i;
                    switch (result.result.Parent.Processes[i - 1].Shori) {
                        case "":
                            readOnly5(id1, id2, id3, id4, id5);
                            break;
                        case "書き込み":
                            readOnly2(id1, id2, id3, id4, id5);
                            placeholder3(id1, id2, id3, "シート名", "セル番地", "値");
                            value3(
                                i
                                , result.result.Parent.Processes[i - 1].Arg1
                                , result.result.Parent.Processes[i - 1].Arg2
                                , result.result.Parent.Processes[i - 1].Arg3
                            )
                            break;
                        default:
                            break;
                    }
                    $('#Processes').append(result.result.PartialView.replace(/Count/g, i + 1));
                }
                console.log('成功');
            },
            error: function () {
                console.log('失敗');
            }
        })
    });
});

function value1(count, arg1) {
    document.getElementById('argument1_' + count).value = arg1;
}

function value2(count, arg1, arg2) {
    document.getElementById('argument1_' + count).value = arg1;
    document.getElementById('argument2_' + count).value = arg2;
}

function value3(count, arg1, arg2, arg3) {
    document.getElementById('argument1_' + count).value = arg1;
    document.getElementById('argument2_' + count).value = arg2;
    document.getElementById('argument3_' + count).value = arg3;
}

function value4(count, arg1, arg2, arg3, arg4) {
    document.getElementById('argument1_' + count).value = arg1;
    document.getElementById('argument2_' + count).value = arg2;
    document.getElementById('argument3_' + count).value = arg3;
    document.getElementById('argument4_' + count).value = arg4;
}

function value5(count, arg1, arg2, arg3, arg4, arg5) {
    document.getElementById('argument1_' + count).value = arg1;
    document.getElementById('argument2_' + count).value = arg2;
    document.getElementById('argument3_' + count).value = arg3;
    document.getElementById('argument4_' + count).value = arg4;
    document.getElementById('argument5_' + count).value = arg5;
}