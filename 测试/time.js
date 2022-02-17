setInterval(main,1)
let musicTimeMinute,musicTimeSecond=0
let musicTimeSecondText,musicTimeMinuteText="" //文本形式显示秒和分
let newTimeWithOffset,offsetTime=0
let nowTimeLabel=""
let maxTimeLabel=""
function main() {
    let b=document.getElementById('audio')
    nowTimeLabel=timeConvert(b.currentTime)
    maxTimeLabel=timeConvert(b.duration)

    function timeConvert(time){
        newTimeWithOffset=time+offsetTime
        musicTimeMinute=Math.trunc(newTimeWithOffset/60) //返回分钟整数部分
        musicTimeSecond=Number(newTimeWithOffset%60).toFixed(2) //秒

        if (musicTimeSecond<10){ //如果秒小于两位数自动补0
            musicTimeSecondText="0"+musicTimeSecond
        }else{
            musicTimeSecondText=musicTimeSecond
        }

        if (musicTimeMinute<10){
            musicTimeMinuteText="0"+musicTimeMinute //如果分钟小于两位数自动补0
        }else{
            musicTimeMinuteText=musicTimeMinute
        }

        return musicTimeMinuteText+":"+musicTimeSecondText
    }
}

