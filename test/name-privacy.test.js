import { test } from 'node:test';
import assert from 'node:assert/strict';
import { loadFirstNamesOnly } from '../dist/config.js';
import { Conversations } from '../dist/conversations.js';

function conversations(firstNamesOnly) {
  const graph = {
    async get(address) {
      if (new URL(address).pathname.endsWith('/me/chats')) {
        return { value: [{
          id: '19:example', chatType: 'oneOnOne', topic: null,
          members: [{ displayName: '  Tayam   El Jebari ' }, { displayName: 'Tim\tMeijer' }],
        }] };
      }
      return { value: [{
        id: 'message',
        createdDateTime: '2026-09-13T10:00:00Z',
        lastModifiedDateTime: '2026-09-13T10:00:00Z',
        from: { user: { displayName: 'Tayam El Jebari' } },
        body: {
          contentType: 'html',
          content: '<p>Hi <at id="0">Tim Meijer</at> and <at id="1">Alex Smith</at></p>',
        },
        mentions: [{ id: 0, mentionText: 'Tim Meijer' }],
      }] };
    },
  };
  return firstNamesOnly === undefined ? new Conversations(graph) : new Conversations(graph, firstNamesOnly);
}

test('first-name privacy defaults on and requires explicit false to disable', () => {
  assert.equal(loadFirstNamesOnly({}), true);
  assert.equal(loadFirstNamesOnly({ TEAMS_MCP_FIRST_NAMES_ONLY: 'true' }), true);
  assert.equal(loadFirstNamesOnly({ TEAMS_MCP_FIRST_NAMES_ONLY: 'typo' }), true);
  assert.equal(loadFirstNamesOnly({ TEAMS_MCP_FIRST_NAMES_ONLY: ' false ' }), false);
});

test('default privacy covers members, generated chat names, senders and mentions in both formats', async () => {
  for (const response_format of ['concise', 'detailed']) {
    const service = conversations();
    const listed = await service.list({ response_format });
    assert.deepEqual(listed.conversations[0].members, ['Tayam', 'Tim']);
    assert.equal(listed.conversations[0].name, 'Tayam, Tim');
    const result = await service.read({ chat: 'Tayam, Tim', response_format });
    assert.equal(result.conversation.name, 'Tayam, Tim');
    assert.equal(result.messages[0].sender, 'Tayam');
    assert.equal(result.messages[0].quotedBody, 'Hi @Tim and @Unknown');
    assert.doesNotMatch(JSON.stringify({ listed, result }), /El Jebari|Meijer|Smith/);
  }
});

test('disabling privacy restores full display names and unresolved mention text', async () => {
  const service = conversations(false);
  const listed = await service.list();
  assert.deepEqual(listed.conversations[0].members, ['Tayam   El Jebari', 'Tim\tMeijer']);
  const result = await service.read({ chat: '19:example' });
  assert.equal(result.messages[0].sender, 'Tayam El Jebari');
  assert.equal(result.messages[0].quotedBody, 'Hi @Tim Meijer and Alex Smith');
});