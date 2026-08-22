/**
 * UpdateMembers.gs — updateAfterUndelete_(), the retry wrapper for the Users.update
 * call right after Users.undelete. A real restore surfaced this: the account came
 * back from Users.undelete but stayed suspended, because the immediate Users.update
 * 404'd (Google hadn't finished propagating the undelete yet) and executeWithRetry
 * doesn't treat 404 as transient — so the un-suspend silently failed and only the
 * NEXT scheduled run's ordinary update caught it. This wrapper retries specifically
 * on "Resource Not Found" so restore + un-suspend land in the same run.
 *
 * Run: npm test
 */
const path = require('path');
const { loadModule, makeLogger, makeChecker } = require('./helpers/apps-script');

const MODULE = path.join(__dirname, '..', 'src', 'accounts-and-groups', 'UpdateMembers.gs');
const { section, check, done } = makeChecker();

/** A fake AdminDirectory.Users.update that fails with "Resource Not Found" N times
 * before succeeding (or fails every time, if failCount >= Infinity-ish). */
function makeFlakyUsersUpdate(failCount) {
  let calls = 0;
  return {
    Users: {
      update: (updates, email) => {
        calls++;
        if (calls <= failCount) {
          throw new Error('Resource Not Found: userKey');
        }
        return { primaryEmail: email, updates: updates, callsToSucceed: calls };
      }
    }
  };
}

function load(AdminDirectory, sleepLog) {
  return loadModule(MODULE, {
    Logger: makeLogger().logger,
    AdminDirectory: AdminDirectory,
    Utilities: { sleep: ms => sleepLog.push(ms) }
  }, ['updateAfterUndelete_']);
}

// ---------------------------------------------------------------------------
section('updateAfterUndelete_ — retries "Resource Not Found" right after undelete');
{
  {
    const sleeps = [];
    const result = load(makeFlakyUsersUpdate(0), sleeps).updateAfterUndelete_('a@x.org', { suspended: false });
    check('succeeds immediately when the account is already queryable',
      result.callsToSucceed, 1);
    check('...and never sleeps', sleeps.length, 0);
  }

  {
    const sleeps = [];
    const result = load(makeFlakyUsersUpdate(2), sleeps).updateAfterUndelete_('a@x.org', { suspended: false });
    check('succeeds on the 3rd attempt after two 404s',
      result.callsToSucceed, 3);
    check('slept twice, with increasing delays',
      sleeps.length, 2);
    check('delays actually increase',
      sleeps[1] > sleeps[0], true);
  }

  {
    const sleeps = [];
    let threw = null;
    try {
      load(makeFlakyUsersUpdate(999), sleeps).updateAfterUndelete_('a@x.org', { suspended: false });
    } catch (e) {
      threw = e.message;
    }
    check('gives up and throws after exhausting all attempts',
      threw !== null, true);
  }

  {
    const sleeps = [];
    let threw = null;
    const AdminDirectory = {
      Users: { update: () => { throw new Error('Invalid Ownership'); } }
    };
    try {
      load(AdminDirectory, sleeps).updateAfterUndelete_('a@x.org', { suspended: false });
    } catch (e) {
      threw = e.message;
    }
    check('a non-"Resource Not Found" error is NOT retried — fails immediately',
      threw, 'Invalid Ownership');
    check('...and never sleeps',
      sleeps.length, 0);
  }
}

done();
