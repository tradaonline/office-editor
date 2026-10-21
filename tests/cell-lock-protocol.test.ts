import assert from "node:assert";
import { test } from "node:test";
import { EditorServer } from "../utils/editor/server";
import { MockSocket } from "../utils/editor/socket";

test("cell lock WebSocket protocol echoes user.id + indexUser to prevent SDK indexOf crash", async () => {
  const server = new EditorServer();
  const socket = new MockSocket({ debug: false });
  server.handleConnect({ socket });

  const clientMessages: any[] = [];
  socket.on("message", (msg: any) => {
    clientMessages.push(msg);
  });

  // Client performs auth
  await server.handleMessage({ type: "auth" });
  const authMsg = clientMessages.find((m) => m.type === "auth");
  assert.ok(authMsg, "Server must respond with auth message");
  assert.strictEqual(authMsg.result, 1);
  assert.strictEqual(authMsg.indexUser, 1);
  assert.strictEqual(authMsg.participants[0].idOriginal, "uid");
  assert.strictEqual(authMsg.participants[0].id, "uid1");

  // Client requests cell lock during editing
  clientMessages.length = 0;
  await server.handleMessage({ type: "getLock", block: "0" });

  const getLockMsg = clientMessages.find((m) => m.type === "getLock");
  assert.ok(getLockMsg, "Server must reply with getLock");
  assert.ok(getLockMsg.locks && getLockMsg.locks["0"]);

  // Exact protocol identity: must be user.id + indexUser (uid1), NOT raw user.id (uid)
  assert.strictEqual(
    getLockMsg.locks["0"].user,
    "uid1",
    "Lock user identity must echo user.id + indexUser per ONLYOFFICE cell SDK expectation",
  );
  assert.notStrictEqual(
    getLockMsg.locks["0"].user,
    "uid",
    "Raw user.id causes SDK y.user !== this.YLb, triggering remote lock AFc -> bEi indexOf exception",
  );

  const releaseLockMsg = clientMessages.find((m) => m.type === "releaseLock");
  assert.ok(releaseLockMsg, "Server must reply with releaseLock");
  assert.strictEqual(releaseLockMsg.locks["0"].user, "uid1");

  // Simulate ONLYOFFICE cell SDK lock state calculation (sdk-all-min.js lines 258, 268)
  const sdkYLb = authMsg.participants[0].idOriginal + authMsg.indexUser;
  const lockState = getLockMsg.locks["0"].user === sdkYLb ? 2 : 3;
  assert.strictEqual(
    lockState,
    2,
    "SDK lock state must be 2 (current user). State 3 triggers remote-lock branch and crashes on undefined sheetId",
  );
});

test("cell lock protocol identity preserves custom user set via setUser", async () => {
  const server = new EditorServer();
  server.setUser({ id: "p0-tester", name: "Người thử P0" });

  const socket = new MockSocket({ debug: false });
  server.handleConnect({ socket });

  const clientMessages: any[] = [];
  socket.on("message", (msg: any) => {
    clientMessages.push(msg);
  });

  await server.handleMessage({ type: "auth" });
  const authMsg = clientMessages.find((m) => m.type === "auth");
  assert.strictEqual(authMsg.participants[0].idOriginal, "p0-tester");
  assert.strictEqual(authMsg.participants[0].id, "p0-tester1");

  clientMessages.length = 0;
  await server.handleMessage({ type: "getLock", block: "cell-A1" });

  const getLockMsg = clientMessages.find((m) => m.type === "getLock");
  assert.strictEqual(getLockMsg.locks["cell-A1"].user, "p0-tester1");

  const releaseLockMsg = clientMessages.find((m) => m.type === "releaseLock");
  assert.strictEqual(releaseLockMsg.locks["cell-A1"].user, "p0-tester1");
});
