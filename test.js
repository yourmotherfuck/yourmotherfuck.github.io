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




    // 切歌的配套变量
    image_music = document.getElementById("image_music")
    music_src = document.getElementById("music_src")
    player_song_name = document.getElementById("player_song_name")
    module3_sign = document.getElementById("module3_sign")
    // 这里留个位置给歌词用



    test_button = document.getElementById("test_button")
    test_text = "看到这个文本说明你成功了"
    demo = document.getElementById("demo")

    // 加载歌曲信息
    first_time = true
    music_change()

    // 加载歌曲列表
    song_list = document.getElementById("song_list")
    load_songlist()
}
    // 这里开始弄数据
    var song_key_list = [ "花海-周杰伦" , "七里香-周杰伦" , "我怀念的-孙燕姿" , "关键词-林俊杰" , "天外来物-薛之谦" ]
    var song_value_list = {
        "花海-周杰伦" : {
            "name" : "花海-周杰伦" ,
            "picture_src" : "music_picture/花海-周杰伦.jpg"  ,
            "music_src" : "music/周杰伦 - 花海.mp3"  ,
            "songword_src" : "",
            "link" : "https://www.bilibili.com/video/BV1iJ411e7ZH?from=search&seid=139711800084340431&spm_id_from=333.337.0.0"
        },
        "七里香-周杰伦" : {
            "name" : "七里香-周杰伦" ,
            "picture_src" : "music_picture/七里香-周杰伦.jpg" ,
            "music_src" : "music/周杰伦 - 七里香.mp3" ,
            "songword_src" : "",
            "link" : "https://www.bilibili.com/video/BV1qD4y1U7fs?from=search&seid=839890076452782897&spm_id_from=333.337.0.0"
        },
        "我怀念的-孙燕姿" : {
            "name" : "我怀念的-孙燕姿" ,
            "picture_src" : "music_picture/逆光-孙燕姿.jpg" ,
            "music_src" : "music/孙燕姿 - 我怀念的.mp3" ,
            "songword_src" : "",
            "link" : "https://www.bilibili.com/video/BV1gZ4y147Z5?from=search&seid=16826540298733889715&spm_id_from=333.337.0.0"
        },
        "关键词-林俊杰" : {
            "name" : "关键词-林俊杰" ,
            "picture_src" : "music_picture/关键词-林俊杰.png" ,
            "music_src" : "music/关键词 - 林俊杰.mp3" ,
            "songword_src" : "music/",
            "link" : "https://www.bilibili.com/video/BV1Ks41197FY?from=search&seid=18108239941964954963&spm_id_from=333.337.0.0"
        },
        "天外来物-薛之谦" : {
            "name" : "天外来物-薛之谦" ,
            "picture_src" : "music_picture/天外来物-薛之谦.jpg" ,
            "music_src" : "music/天外来物 - 薛之谦.mp3" ,
            "songword_src" : "music/天外来物 - 薛之谦.lrc" ,
            "link" : "https://www.bilibili.com/video/BV11h411Q7vz?from=search&seid=4404551429510665856&spm_id_from=333.337.0.0"
        }

        // 歌曲添加模板
        // "name" : "" ,
        // "picture_src" : "music_picture/" ,
        // "music_src" : "music/" ,
        // "songword_src" : "music/" ,
        // "link" : ""
    }

    function button(){
        demo.innerHTML =  document.getElementById("link").href = song_value_list[ song_key_list[nowsong_index] ]["link"]

    }
    //进入先加载歌单列表
    function load_songlist() {
        let song_index
        for ( song_index=0; song_index < music_count;song_index++){
            song_list.innerHTML +=  "<li onclick=\"nowsong_index =" + song_index +
                "; music_change();\"><a href=\"#\">" + song_key_list[song_index] +
                "</a></li>"
        }
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
            music_change()
        }else if ( play_order === "shuffle" ){
            let temp = Math.floor( Math.random() * music_count )
            while (temp === nowsong_index){
                temp = Math.floor( Math.random() * music_count )
            }
            nowsong_index = temp
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
        }else if ( play_order === "shuffle" ){
            let temp = Math.floor( Math.random() * music_count )
            while (temp === nowsong_index){
                temp = Math.floor( Math.random() * music_count )
            }
            nowsong_index = temp
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
            next()
        }
    }
    function totaltime() {
        audio.oncanplay = function () {
            // 这个是总时长（秒为单位）
            let second = Math.floor( audio.duration )
            let total_minute = Math.floor( second / 60 )
            let total_second = second % 60

            if ( total_second < 10 ){
                total_second = "0" + total_second
            }
            document.getElementById("total_time").innerHTML = total_minute + ":" + total_second
        }

    }
    
    
    function music_change() {
        pause()
        image_music.src = song_value_list[song_key_list[nowsong_index]]["picture_src"]
        audio.src = song_value_list[song_key_list[nowsong_index]]["music_src"]
        player_song_name.innerHTML = "当前播放：" + song_value_list[song_key_list[nowsong_index]]["name"]
        module3_sign.innerHTML = song_value_list[song_key_list[nowsong_index]]["name"]
        // 留一行给歌词文件
        document.getElementById("music_link").href = song_value_list[ song_key_list[nowsong_index] ]["link"]
        audio.load()
        totaltime()

        // 检测为第一次进不会自动播放
        if ( first_time ){
            pause()
            first_time = false
        }else{
            play()
        }
        // 图片自转重置
        // image_music.style = "transform: rotate( 0deg );"
    }
