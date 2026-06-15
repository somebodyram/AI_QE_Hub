using { exceldata as my } from '../db/schema.cds';

@path : '/service/empservice'
service empservice {
  @cds.redirection.target
  @odata.draft.enabled
  
  entity TestResults as projection on my.TestResults;
}

annotate empservice with @requires : [
  'authenticated-user'
];