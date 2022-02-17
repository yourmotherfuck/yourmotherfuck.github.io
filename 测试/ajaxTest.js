function main() {
    setInterval(timeGet,1)
    let c=0
    function timeGet(){
        document.getElementById('time').innerHTML = maxTimeLabel +" "+Number(c+1)
        for (let x=0;x<document.getElementById('lyricContainer').children.length;x++){
            let timeABC = document.getElementById('lyricContainer').children[x].innerHTML.match(timeLabel)
            if (timeABC!=null){
                c=x
                try {
                    document.getElementById('lyricContainer').children[c].style.color="Red"
                    document.getElementById('lyricContainer').children[c-1].style.color="Black"
                }catch (e) {

                }

            }
        }
    }

    let xmlGet = new XMLHttpRequest()
    xmlGet.onreadystatechange=function(){
        if (xmlGet.readyState===4 && xmlGet.status===200){
            let ajaxtext = xmlGet.responseText
            let ab = ajaxtext.split('\n')
            for (let x =0;x<ab.length;x++){
                let sb = ab[x].replace(/.*/,"<div>"+"$&"+"</div>")
                document.getElementById('lyricContainer').innerHTML+=sb
            }
        }
    }
    xmlGet.open("get","花海.lrc",true)
    xmlGet.send()
}
main()