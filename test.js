window.onload = function() {
    pause_play = document.getElementById("pause_play")
    love_variable = document.getElementById("love_button")
    repeat_variable = document.getElementById("repeat")
    previous_variable = document.getElementById("previous")
    play_order = "repeat"

    volume_variable = document.getElementById("volume")
    volume_index = 0
    volume_src = ["player_image/volume-up.png","player_image/volume-down.png","player_image/mute.png"]

    audio = document.getElementById("audio")

    // 这里开始弄数据
    song_list = [  ]

    test_button = document.getElementById("test_button")
    test_text = "看到这个文本说明你成功了"
    demo = document.getElementById("demo")
    totaltime()
}

    function button(){
        let x = document.getElementById("audio")
        x.volume = 0.0

    }


    // 播放或者暂停按钮点击后反应
    function main_button() {
        if (pause_play.getAttribute("src") === "player_image/play.png") {
            pause()
        } else if (pause_play.getAttribute("src") === "player_image/pause.png") {
            play()
        }
    }
    //开始和暂停
    function pause() {
        pause_play.setAttribute("src","player_image/pause.png")
        document.getElementById("image_music").style.animationPlayState = "paused"
        audio.pause()
    }
    function play() {
        pause_play.setAttribute("src","player_image/play.png")
        document.getElementById("image_music").style.animationPlayState = "running"
        audio.play()

    }

    //随机播放和顺序播放切换模式
    function repeat() {
        if ( repeat_variable.getAttribute("src") === "player_image/repeat.png" ){
            repeat_variable.setAttribute("src","player_image/shuffle.png")
            window.play_order = "shuffle"
        }else if ( repeat_variable.getAttribute("src") === "player_image/shuffle.png" ){
            repeat_variable.setAttribute("src","player_image/repeat.png")
            window.play_order = "repeat"
        }
    }

    //这里设置随机和顺序状态下的上一曲和下一曲(第一模块为顺序播放)
    function previous() {
        if ( play_order === "repeat" ){
            demo.innerHTML = "repeat"
            alert("现在是顺序播放")
        }else if ( play_order === "shuffle" ){
            demo.innerHTML = "shuffle"
            alert("现在是随机播放")
        }
    }
    function next() {
        if ( play_order === "repeat" ){
            demo.innerHTML = "repeat"
            alert("现在是顺序播放")
        }else if ( play_order === "shuffle" ){
            demo.innerHTML = "shuffle"
            alert("现在是随机播放")
        }
    }

    function volume(){
        window.volume_index += 1
        if ( volume_index === 3 ){
            window.volume_index =0
        }
        if (volume_index === 0){
            audio.volume = 1.0
       }else if (volume_index === 1){
            audio.volume = 0.5
        }else if (volume_index === 2){
            audio.volume = 0.0
        }
        volume_variable.src = volume_src[volume_index]
    }


    function love_button(){
        if ( love_variable.getAttribute("src") === "player_image/love.png" ){
            love_variable.src = "player_image/love - 副本.png"
        }else if (love_variable.getAttribute("src") === "player_image/love - 副本.png"){
            love_variable.src = "player_image/love.png"
        }
    }
    // 当前时间加上总时长
    function nowtime(event) {
        let now_minute = Math.floor( event.currentTime / 60 )
        let now_second = Math.floor( event.currentTime) % 60


        if ( now_second < 10 ){
            now_second = "0" + now_second
        }

        document.getElementById("now_time").innerHTML = now_minute.toString() + ":" + now_second.toString()
        if ( event.currentTime >= audio.duration ){
            pause()
        }
    }
    function totaltime() {
        let total_minute = Math.floor(Math.floor(audio.duration) / 60)
        let total_second = Math.floor( audio.duration) % 60

        if ( total_second < 10 ){
            total_second = "0" + total_second
        }
        document.getElementById("total_time").innerHTML = total_minute + ":" + total_second
    }
