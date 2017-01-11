/// <reference path="../typings/globals/jquery/index.d.ts" />

declare var _spPageContextInfo:any;

/**
 * This is a class that assists in tracking users  
 */
class Checklist{
    /**
     * @param listTitle Title of the List that stores
     * @param listDescription List description
     */
    constructor(public listTitle:string, public listDescription:string = ""){
    }

    /**
     * Ensures that list is created
     */
    ensureList():JQueryPromise<void>{
        let that = this;
        return this.listExists().then(function(listExists){
            if (listExists){
                console.log(that.listTitle + ' already exists. do nothng');
            }
            else{
                console.log(that.listTitle + " being created...");
                that.getFormDigest().then(function(digest){
                    that.createList(digest).then(function(data){
                        console.log(that.listTitle + " created");
                    })
                });
            }
        });
    }
    /**
     * Create Check List
     * 
     * @param digest Formdigest, retrieved using getFormDigest method
     */
    createList(digest):JQueryPromise<any>{
        let settings:JQueryAjaxSettings;
        let data = JSON.stringify({ 
                 '__metadata': { 'type': 'SP.List' },
                 'AllowContentTypes': false,
                 'BaseTemplate': 100,
                 'ContentTypesEnabled': true,
                 'Description': this.listDescription,
                 'Title': this.listTitle 
                });
        settings = {
            url: _spPageContextInfo.webServerRelativeUrl + "/_api/web/lists",
            method: "POST",
            data: data,
            dataType: "JSON",
            contentType: "application/json;odata=verbose",
            headers: {  
                "accept": "application/json;odata=verbose",
                "X-RequestDigest":digest,
            },
        };

        return  jQuery.ajax(settings).then(function(data, textStatus){
            return data
        }).fail(function(err){
            console.log(err);
        });
    }

    /**
     * Check if list if exists
     */
    listExists():JQueryPromise<boolean>{
        let settings:JQueryAjaxSettings;
        settings = {
            url: _spPageContextInfo.webServerRelativeUrl + "/_api/Web/Lists?$filter=title eq '"+this.listTitle+"'",
            method: "GET",
            headers: { "accept": "application/json; odata=verbose" },
        };

        return jQuery.ajax(settings).then(function(data:any, textStatus:string){
            if (data.d.results.length === 0){
                return false;
            }
            else{
                return true;
            }
        }).fail(function(err){
            console.log(err);
            return err;
        })
    }

    /**
     * Get Form digest 
     */
    getFormDigest():JQueryPromise<string>{
        return $.ajax({
                url: _spPageContextInfo.webServerRelativeUrl + "/_api/contextinfo",
                method: "POST",
                headers: { "Accept": "application/json; odata=verbose" }
            }).then(function(data){
                return data.d.GetContextWebInformation.FormDigestValue;
        });
    }

    addCheckItem(){
        
    }

    /**
     * Check if list item if exists
     */
    checkItemExists():JQueryPromise<boolean>{
        let settings:JQueryAjaxSettings;
        settings = {
            url: _spPageContextInfo.webServerRelativeUrl + "/_api/Web/Lists/getbytitle('"+this.listTitle+"')/items?$filter=AuthorId eq " + _spPageContextInfo.userId,
            method: "GET",
            headers: { "accept": "application/json; odata=verbose" },
        };

        return jQuery.ajax(settings).then(function(data:any, textStatus:string){
            if (data.d.results.length === 0){
                return false;
            }
            else{
                return true;
            }
        }).fail(function(err){
            console.log(err);
            return err;
        })
    }

    /**
     * Create Check List Item
     * 
     * @param digest Formdigest, retrieved using getFormDigest method
     */
    createCheckItem(digest, ListItemEntityTypeFullName):JQueryPromise<any>{
        let settings:JQueryAjaxSettings;
        let data = JSON.stringify({ 
                 '__metadata': { 'type': ListItemEntityTypeFullName }, 
                'Title': JSON.stringify(_spPageContextInfo.userId)
                });
        settings = {
            url: _spPageContextInfo.webServerRelativeUrl + "/_api/web/lists/getbytitle('"+this.listTitle+"')/items",
            method: "POST",
            data: data,
            dataType: "JSON",
            contentType: "application/json;odata=verbose",
            headers: {  
                "accept": "application/json;odata=verbose",
                "X-RequestDigest":digest,
            },
        };

        return  jQuery.ajax(settings).then(function(data, textStatus){
            return data
        }).fail(function(err){
            console.log(err);
        });
    }

        /**
     * Ensures that check list item is created
     */
    ensureCheckListItem():JQueryPromise<void>{
        let that = this;
        return this.checkItemExists().then(function(checkListItemExists){
            if (checkListItemExists){
                console.log('check list item already exists. do nothng');
            }
            else{
                
                console.log("checklist item being created...");
                that.getFormDigest().then(function(digest){
                    that.getListEntityFullName().then(function(ListEntityFullName){
                         that.createCheckItem(digest, ListEntityFullName ).then(function(data){
                            console.log("check item created");
                        })
                    });
                });
            }
        });
    }

    getListEntityFullName(){
        let settings:JQueryAjaxSettings = {
            url:  _spPageContextInfo.webServerRelativeUrl +
            "/_api/web/lists/getbytitle('" + this.listTitle + "')/ListItemEntityTypeFullName",
            method:"GET",
            headers: { "accept": "application/json; odata=verbose" }
        }
        return jQuery.ajax(settings).then(function(data){
            return data.d.ListItemEntityTypeFullName;
        });
    }

    removeCheckItem(){
        //TODO: remove check item
    }
}  

jQuery(function(){
    //Example of use
    //let checklist:Checklist = new Checklist('lala','descr');
   // checklist.ensureList().then(function(){
        //checklist.checkItemExists();
        //checklist.ensureCheckListItem();
    //})
});