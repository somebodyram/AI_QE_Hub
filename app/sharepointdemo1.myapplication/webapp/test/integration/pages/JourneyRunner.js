sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"sharepointdemo1/myapplication/test/integration/pages/empdataList",
	"sharepointdemo1/myapplication/test/integration/pages/empdataObjectPage"
], function (JourneyRunner, empdataList, empdataObjectPage) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('sharepointdemo1/myapplication') + '/test/flpSandbox.html#sharepointdemo1myapplication-tile',
        pages: {
			onTheempdataList: empdataList,
			onTheempdataObjectPage: empdataObjectPage
        },
        async: true
    });

    return runner;
});

