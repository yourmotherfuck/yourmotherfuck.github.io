window.onload = function () {
    song_index = 0
    song_list = { "one" : [0,1,2,3,4] }
    
    marginbox = document.getElementById("marginbox")
    marginbox.onmousedown = function (event) {
        let offsetx = event.offsetX
        if ( offsetx < 0 ){
            offsetx = 0
        }else if (offsetx > 400){
            offsetx = 400
        }
        test2.innerHTML = offsetx
        box.style.marginLeft = offsetx + "px"

    }
    box.onload = function (e) {
        e.stopPropagation()
        alert("success")
    }

}

function test2_function() {
    // 文本设置
    let test2 = document.getElementById("test2");
    reader = new FileReader()
    reader.readAsText( document.getElementById("fileinput").files[0] )
    reader.onload = function () {
        test2.innerText =  this.result

    // 设置歌词

    }




    // 图片设置
    let test1 = document.getElementById("test1");
    let a = 123

    //点击按钮播放按钮
    // document.getElementById("audio").play()
}
function test10086() {
    document.getElementById("test2").innerHTML = 123
}
