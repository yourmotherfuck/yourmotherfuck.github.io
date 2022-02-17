window.onload = function () {
    songword = document.getElementById("songword")
    fortest = document.getElementById("fottest")
    fortest1 = document.getElementById("fottest1")
    Array = []
    a=[]
    audio = document.getElementById("audio")
    setInterval(wordtest,1)
}
function test1() {
    every_word = songword.innerHTML.replace(/\[[0-9][0-9]:[0-9][0-9].[0-9][0-9]\]/g,"")
    every_time = songword.innerHTML.match(/[0-9][0-9]:[0-9][0-9].[0-9][0-9]/g)
    regex = /[\u4e00-\u9fa5\u0800-\u4e00A-Za-z0-9]/
    if ( ! regex.test( every_word[3] ) ){

    }
    for ( i=0;i<every_word.length;i++ ){
        if ( ! regex.test( every_word[i+1] ) && i !== every_word.length -1 ){
            a.push( every_word[i] + every_word[Number(i)+1] )
            alert(123)
        }else if ( ! regex.test( every_word[i] ) ){

        }else {
            a.push( every_word[i] )
        }
    }
    fortest.innerHTML = a
    // for ( i=0;i<nospaceword.length;i++ ){
    //
    //
    //     Array.push([every_word[i],every_time[i],every_time[i+1]])
    // }

    // fortest1.innerHTML =  every_time.length
    // for (i in a ){
    //     alert(a[i])
    // }

}
function wordtest() {
    
}