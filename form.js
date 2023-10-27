Office.onReady(function(info) {

    if (info.isOfficeReady) {

        function moveSelectedItemToDeletedItems() {
            var item = Office.context.mailbox.item;
            var itemId = item.itemId;
            

            easyEws.moveItem(itemId, "deleteditems", function() {
                console.log("Spostamento riuscito");
            }, function(error) {
                console.log("Errore durante lo spostamento: " + error);
            });
        }

        document.getElementById("moveToDeletedItemsButton").addEventListener("click", moveSelectedItemToDeletedItems);
    }
});