/**
 * SyncOrgPaths.gs — OU path segments follow CONFIG.WING, not a literal 'CA-'.
 *
 * Before 1.2.0 the derived path was '<parent>/CA-<unit>' and the deactivation
 * scan only looked at paths starting '/CA-001', so on any other wing new orgs
 * were provisioned as CA-named OUs and deactivations were never reported.
 * All data is synthetic.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'SyncOrgPaths.gs');
const { section, check, done } = makeChecker();

// Organization.txt: ORGID, Region, Wing, Unit, Parent, Name, Type, Chartered, Status, Scope
const org = (id, wing, unit, parent, status, scope) =>
  [id, 'PCR', wing, unit, parent, 'Unit ' + unit, 'X', '', status, scope || 'UNIT'];

function run(wing, orgRows, pathRows) {
  const created = [];
  const written = [];
  const files = { Organization: orgRows, OrgPaths: pathRows };
  const m = loadModule(MODULE, {
    Logger: makeLogger().logger,
    CONFIG: { WING: wing, ORG_LABEL: wing + 'WG', CAPWATCH_DATA_FOLDER_ID: 'f' },
    parseFile: (n) => files[n],
    AdminDirectory: { Orgunits: { insert: (ou) => created.push(ou.parentOrgUnitPath + '/' + ou.name) } },
    MailApp: { sendEmail: () => {} },
    Session: { getScriptTimeZone: () => 'UTC' },
    Utilities: { formatDate: () => 'today' },
    DriveApp: {
      getFolderById: () => ({
        getFilesByName: () => {
          let done_ = false;
          return { hasNext: () => !done_, next: () => { done_ = true; return {
            getBlob: () => ({ getDataAsString: () => '' }),
            setContent: (c) => written.push(c) }; } };
        }
      })
    },
  }, ['syncOrgPaths']);
  return { result: m.syncOrgPaths(), created, written };
}

section('1. A non-California wing derives its own prefix');
{
  const r = run('OR',
    [org('1', 'OR', '001', '0', 'ACTIVE', 'WING'), org('2', 'OR', '123', '1', 'ACTIVE')],
    [['1', '/OR-001']]);
  check('one org added', r.result.added, 1);
  check('OU is OR-named, not CA-named', r.created, ['/OR-001/OR-123']);
  check('OrgPaths.txt gets the OR path', r.written.join('').trim(), '2,/OR-001/OR-123');
}

section('2. Deactivations are detected outside California');
{
  const r = run('OR',
    [org('1', 'OR', '001', '0', 'ACTIVE', 'WING')],
    [['1', '/OR-001'], ['9', '/OR-001/OR-999']]);
  check('a mapped org with no active match is reported', r.result.deactivated, 1);
}

section('3. Other wings\' paths are left alone');
{
  const r = run('OR',
    [org('1', 'OR', '001', '0', 'ACTIVE', 'WING')],
    [['1', '/OR-001'], ['500', '/CA-001/CA-070']]);
  check('a CA path in an OR tenant is not flagged', r.result.deactivated, 0);
}

section('4. California is unchanged');
{
  const r = run('CA',
    [org('1', 'CA', '001', '0', 'ACTIVE', 'WING'), org('2', 'CA', '404', '1', 'ACTIVE')],
    [['1', '/CA-001']]);
  check('CA org added under CA-', r.created, ['/CA-001/CA-404']);
}

section('5. An unset wing refuses rather than guessing');
{
  let err = '';
  try { run('', [], []); } catch (e) { err = e.message; }
  check('throws naming TENANT_WING', /TENANT_WING/.test(err), true);
}

done();
