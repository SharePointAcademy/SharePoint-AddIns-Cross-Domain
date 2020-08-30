'use strict';

var hostweburl;
var appweburl;

$(document).ready(function () {
    hostweburl =
        decodeURIComponent(
            getQueryStringParameter("SPHostUrl")
        );
    appweburl =
        decodeURIComponent(
            getQueryStringParameter("SPAppWebUrl")
        );

    console.log("hostweburl", hostweburl);
    console.log("appweburl", appweburl);

    var scriptbase = hostweburl + "/_layouts/15/";
    $.getScript(scriptbase + "SP.RequestExecutor.js", execCrossDomainRequest);
});

function execCrossDomainRequest() {
    var executor = new SP.RequestExecutor(appweburl);

    executor.executeAsync(
        {
            url: appweburl + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('Product')/items?@target='" + hostweburl + "'&$top=10&$orderby=ID desc",
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" },
            success: successHandler,
            error: errorHandler
        }
    );
}

function successHandler(data) {
    var jsonObject = JSON.parse(data.body);
    var productsHTML = "";
    var results = jsonObject.d.results;
    for (var i = 0; i < results.length; i++) {
        productsHTML = productsHTML + "<div>" + results[i].name + "</div><br>";
    }

    $('#products').html(productsHTML);

}

function errorHandler(data, errorCode, errorMessage) {
    document.getElementById("noticias").innerText =
        "Could not complete cross-domain call: " + errorMessage;
}


function getQueryStringParameter(paramToRetrieve) {
    var params =
        document.URL.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] === paramToRetrieve)
            return singleParam[1];
    }
}

function adicionarProduto() {
    //insere um produto
    var data = {
        __metadata: { 'type': 'SP.Data.ProductListItem' },
        name: 'a-SampleData',
        stock: 10,
        price: 2.60
    };

    var urlSiteDev = appweburl + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('Product')/items?@target='" + hostweburl + "'";
    addNewItem(urlSiteDev, data);
}

function addNewItem(restUrl, data) {
    $.ajax({
        url: restUrl,
        type: "POST",
        data: JSON.stringify(data),
        headers: {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "X-HTTP-Method": "POST"
        },
        success: function (data) {
            alert("Produto inserido com sucesso");
            execCrossDomainRequest();

        },
        error: function (error) {
            console.log(JSON.stringify(error));
        }
    });
}

