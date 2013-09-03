$(document).ready(function(){

    // Initialize the jQuery File Upload widget:
    $('#fileupload').fileupload({
        // Uncomment the following to send cross-domain cookies:
        //xhrFields: {withCredentials: true},
        url: '/upload'
    });
    $.get('/ingredients', function(data) {
        store = new Persist.Store('My Application');
        store.get('some_key', function(ok, val) {
            if (val)
                console.log(JSON.parse(val)['test'])
        });
        store.set('some_key', JSON.stringify({'test':1}));
    });
});