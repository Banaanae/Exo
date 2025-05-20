const textarea = document.createElement('textarea')
textarea.id = 'textarea'

textarea.onkeydown = function(e){
    MsgBox('Detected: ' + e.key, 'Key pressed', 'YN')
}

document.querySelector('body').appendChild(textarea)