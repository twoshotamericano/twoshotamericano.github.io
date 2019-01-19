(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            //Apply Bindings
            
            ko.applyBindings(viewModel)

            //Initialize the dataservice
            ds = trelloDataService();

            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#authorise-button-text').text("Authorise");
            $('#deauthorise-button-text').text("Deauthorise");
            $('#load-boards-button-text').text("Load boards!");
                
  
            // Add a click event handler for the highlight button.
            $('#authorise-button').click(ds.authorize);
            $('#deauthorise-button').click(ds.deauthorize);
            $('#load-boards-button').click(ds.loadBoards);
        });
    };

    //Knock-Out

    function beard(colour) {
        return {name:ko.observable(colour)}
    }

    var viewModel = {
        boards: ko.observableArray([]),
        beards: ko.observableArray([new beard('red'),new beard('green')])
    };

  

    //Trello

    var ds = {};

    function trelloDataService() {

        var dataService = {
            authorize: trelloAuthorize,
            deauthorize: trelloDeAuthorize,
            loadBoards: trelloLoadBoards
        };

        return dataService;

    };

    var authenticationSuccess = function () {
        showNotification('Authorise Buton', 'Pressed');
    };

    var authenticationFailure = function () {
        errorHandler(error);
    };

    var loadBoardsSuccess = function (boards) {
        showNotification(boards.length);

        viewModel.boards.removeAll();

        boards.forEach(function (board) {
            console.log('Name', board.name);
            this.boards.push({ name: board.name });
            //console.log(this.boards[0].name)
        }, viewModel);

        showNotification('length',viewModel.boards.length)
    };

    var loadBoardsFailure = function (error) {
        errorHandler(error);
    };

 
    function trelloAuthorize(success, failure) {

        Trello.authorize({
            name: 'First Excel App',
            scope: {
                read: 'true',
                write: 'true'
            },
            expiration: 'never',
            success: authenticationSuccess,
            error: authenticationFailure
        });
        
    };



    function trelloLoadBoards(success, error) {
         Trello.get('/member/me/boards', loadBoardsSuccess, loadBoardsFailure);
    };

    function trelloDeAuthorize(success, failure) {
        showNotification('DeAuthorise Buton', 'Pressed')
        Trello.deauthorize();
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
