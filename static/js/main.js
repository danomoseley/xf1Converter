$(document).ready(function(){
    // Initialize the jQuery File Upload widget:
    $('#fileupload').fileupload({
        url: '/upload'
    }).bind('fileuploaddone', function (e, data) {
        var files = getFiles();
        $.each(data.result.files, function(k, file) {
            console.log(file);
            files[file.url] = file;
        });
        localStorage.setItem("files", JSON.stringify(files))

        updateFileList();
    });
    updateFileList('#fileList');
    processSS();
});

function chooseFile(callback){
    $.colorbox({
        html: getFileList(),
        onComplete: function(){
            $('.file_list .file').on('click', function(){
                callback($(this).data('file'));
                $.colorbox.close();
            });
        }
    })
}

function getFiles(){
    var files = localStorage.getItem("files");
    if (files == null) {
        files = {};
    } else {
        files = JSON.parse(files);
    }
    return files;
}

function getFileList(){
    var list = $('<ul class="file_list">');
    $.each(getFiles(), function(k, v){
        var file = $('<li class="file">').html(k);
        file.data('file', v);
        list.append(file);
    });
    return list;
}

function updateFileList(target){
    var list = getFileList();
    if (typeof target == 'undefined') {
        $('ul.file_list').replaceWith(list);
    } else {
        $(target).append(list);
    }
}

function processSS(){
    var req = {}
    //chooseFile(function(file){
    //    req['agris_csv'] = file.url;
    //});
    chooseFile(function(file){
        if (typeof req['plant_files'] == 'undefined') {
            req['plant_files'] = {};
        }
        req['plant_files']['550'] = file.url;
    });
    $('#process-ss').on('click', function(){
        console.log(req);
    });
}