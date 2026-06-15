sap.ui.define([
    "sap/fe/test/JourneyRunner",
	"sharepointdemo1/custompageapp/test/integration/pages/empdataMain"
], function (JourneyRunner, empdataMain) {
    'use strict';

    var runner = new JourneyRunner({
        launchUrl: sap.ui.require.toUrl('sharepointdemo1/custompageapp') + '/test/flpSandbox.html#sharepointdemo1custompageapp-tile',
        pages: {
			onTheempdataMain: empdataMain
        },
        async: true
    });

    return runner;
});

