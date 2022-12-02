$(function () {
    $("#pageSize").change(function () {
        var size = $("#pageSize option:selected").text();
        $("#form1").submit();
    });
});
