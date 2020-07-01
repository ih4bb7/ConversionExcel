$(function () {
    $('#execute').click(function () {
        var processes = new Array();
        var count = (document.getElementById('Processes').innerHTML.match(/処理内容/g) || []).length;
        for (var i = 1; i < count + 1; i++) {
            var process = {
                Shori: $('#shori_' + i).val(),
                Arg1: $('#argument1_' + i).val(),
                Arg2: $('#argument2_' + i).val(),
                Arg3: $('#argument3_' + i).val(),
                Arg4: $('#argument4_' + i).val(),
                Arg5: $('#argument5_' + i).val(),
            };
            processes.push(process);
        }
        var model = {
            ReadPath: $('#readPath').val(),
            OutputPath: $('#outputPath').val(),
        };
        model.Processes = processes;

        $.ajax({
            contentType: "application/json",
            data: JSON.stringify(model),
            type: "POST",
            url: "/Home/execute_Click",
            success: function (result) {
                alert(result.result);
                console.log('成功');
            },
            error: function () {
                console.log('失敗');
            }
        })
    });
});