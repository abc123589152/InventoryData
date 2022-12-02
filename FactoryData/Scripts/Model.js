$(function () {
     $('#myModel').on('click', 'a', function () {
            $.ajax({
                url: this.href,
                type: 'GET',
                cache: false,
                success: function (result) {
                    $('#Report').html(result);
                }
            });
            return false;
        });
    });
