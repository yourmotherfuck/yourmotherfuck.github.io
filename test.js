window.onload = function() {
    pause_play = document.getElementById("pause_play")
    love_variable = document.getElementById("love_button")
    // 这里用于区分随机和顺序播放
    repeat_variable = document.getElementById("repeat")
    previous_variable = document.getElementById("previous")
    play_order = "repeat"
    //这里具体实现不同情况下的上下曲
    nowsong_index = 0
    music_count = song_key_list.length



    volume_variable = document.getElementById("volume")
    volume_index = 0
    volume_src = ["player_image/volume-up.png","player_image/volume-down.png","player_image/mute.png"]

    audio = document.getElementById("audio")




    // 切歌配套变量
    image_music = document.getElementById("image_music")
    music_src = document.getElementById("music_src")
    player_song_name = document.getElementById("player_song_name")
    module3_sign = document.getElementById("module3_sign")
    // 这里留个位置给歌词用



    test_button = document.getElementById("test_button")
    test_text = "看到这个文本说明你成功了"
    demo = document.getElementById("demo")
    totaltime()
}
    // 这里开始弄数据
    var song_key_list = [ "花海-周杰伦" , "七里香-周杰伦" , "我怀念的-孙燕姿" ]
    var song_value_list = {
        "花海-周杰伦" : {
            "name" : "花海-周杰伦" ,
            "picture_src" : "music_picture/花海-周杰伦.jpg"  ,
            "music_src" : "music/周杰伦 - 花海.mp3"  ,
            "songword_src" : ""
        },
        "七里香-周杰伦" : {
            "name" : "七里香-周杰伦" ,
            "picture_src" : "music_picture/七里香-周杰伦.jpg" ,
            "music_src" : "music/周杰伦 - 七里香.mp3" ,
            "songword_src" : ""
        },
        "我怀念的-孙燕姿" : {
            "name" : "我怀念的-孙燕姿" ,
            "picture_src" : "music_picture/逆光-孙燕姿.jpg" ,
            "music_src" : "music/孙燕姿 - 我怀念的.mp3" ,
            "songword_src" : ""
        }
    }

    function button(){

        demo.innerHTML = Math.floor( Math.random() * music_count )
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
            nowsong_index -= 1
            if (nowsong_index < 0) {
                nowsong_index = song_key_list.length - 1
            }
            pause()
            music_change()
        }else if ( play_order === "shuffle" ){
            let temp = Math.floor( Math.random() * music_count )
            while (temp === nowsong_index){
                temp = Math.floor( Math.random() * music_count )
            }
            nowsong_index = temp
            pause()
            music_change()
        }
    }
    function next() {
        if ( play_order === "repeat" ){
            nowsong_index += 1
            if ( nowsong_index > song_key_list.length -1 ){
                nowsong_index = 0
            }
            music_change()
            pause()
        }else if ( play_order === "shuffle" ){
            let temp = Math.floor( Math.random() * music_count )
            while (temp === nowsong_index){
                temp = Math.floor( Math.random() * music_count )
            }
            nowsong_index = temp
            pause()
            music_change()
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
    
    
    function music_change() {
        image_music.src = song_value_list[song_key_list[nowsong_index]]["picture_src"]
        music_src.src = song_value_list[song_key_list[nowsong_index]]["music_src"]
        player_song_name.innerHTML = "当前播放：" + song_value_list[song_key_list[nowsong_index]]["name"]
        module3_sign.innerHTML = song_value_list[song_key_list[nowsong_index]]["name"]
        // 留一行给歌词文件
        totaltime()
    }
