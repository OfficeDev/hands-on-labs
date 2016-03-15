
$(function () {
    'use strict';

    //render index page in the contect container
    navigation.goTo("index");

    //initialize skype app
    $(".modal").show();

    //side panel animations
    $('.slider-arrow').click(function () {
        var move_width = 390;
        var animation_speed = 300;

        if ($(this).hasClass('show')) {
            $(".slider-arrow, .right-slider").animate({
                right: "+=" + move_width
            }, animation_speed, function () {
                // Animation complete.
            });

            $(this).html('&raquo;').removeClass('show').addClass('hide');

            $(".container").animate({
                width: "-=" + move_width
            }, animation_speed, function () {
                // Animation complete.
            });
        }
        else {
            $(".slider-arrow, .right-slider").animate({
                right: "-=" + move_width
            }, animation_speed, function () {
                // Animation complete.
            });

            $(this).html('&laquo;').removeClass('hide').addClass('show');

            $(".container").animate({
                width: "+=" + move_width
            }, animation_speed, function () {
                // Animation complete.
            });
        }
    });
});