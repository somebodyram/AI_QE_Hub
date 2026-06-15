using empservice as service from '../../srv/service';
annotate service.empdata with @(
    UI.FieldGroup #GeneratedGroup : {
        $Type : 'UI.FieldGroupType',
        Data : [
            {
                $Type : 'UI.DataField',
                Label : 'empid',
                Value : empid,
            },
            {
                $Type : 'UI.DataField',
                Label : 'empname',
                Value : empname,
            },
        ],
    },
    UI.Facets : [
        {
            $Type : 'UI.ReferenceFacet',
            ID : 'GeneratedFacet1',
            Label : 'General Information',
            Target : '@UI.FieldGroup#GeneratedGroup',
        },
    ],
    UI.LineItem : [
        {
            $Type : 'UI.DataField',
            Label : 'empid',
            Value : empid,
        },
        {
            $Type : 'UI.DataField',
            Label : 'empname',
            Value : empname,
        },
    ],
);

