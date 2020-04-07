import * as React from 'react'
import {
  Button,
  ButtonType,
  MessageBar,
  MessageBarType,
  Spinner,
} from 'office-ui-fabric-react'


const CALENDAR_ITEM_CLASSES = [
  'IPM.Schedule.Meeting.Request',
  'IPM.Schedule.Meeting.Resp.Pos',
]

const DATA_HUB_STUB_INTERACTION_FORM_URL = 'http://localhost:3001/interactions/create-stub'

function getUrlToStubInteractionForm({ subject, start, from, to }) {
  const participants = [from, ...to]
  const params = [
    subject && `subject=${subject}`,
    start && `date=${start.toISOString()}`,
    ...participants.map(c => `participant_email=${c.emailAddress}`),
  ].filter(p => p).join('&')

  return new URL(
    '?' + params,
    DATA_HUB_STUB_INTERACTION_FORM_URL)
    .href
}

export default function App({ isOfficeInitialized }) {
  const click = () => setCreated(true)
  console.log(Office.context)

  if (!isOfficeInitialized) {
    return (
      <Spinner />
    );
  }

  const item = Office.context.mailbox.item;

  if (!CALENDAR_ITEM_CLASSES.includes(item.itemClass)) {
    return (
      <MessageBar messageBarType={MessageBarType.info}>
        There is no meeting in the selected email.
      </MessageBar>
    )
  }

  const formUrl = getUrlToStubInteractionForm(item)

  Office.context.mailbox.item

  return (
    <div style={{ margin: '5%' }}>
      <h3>Interaction details</h3>

      <dl>
        {item.subject && <><dt>Subject</dt><dd>{item.subject}</dd></>}
        {item.location && <><dt>Location</dt><dd>{item.location}</dd></>}
        {item.start && <><dt>Date</dt><dd>{item.start.format()}</dd></>}

        <dt>Participants</dt>
        <dd>
          <ul>
            {[item.from, ...item.to.filter(p => p.emailAddress !== item.from.emailAddress)]
              .map(p =>
                <li key={p.emailAddress}>
                  {p.displayName}{p.emailAddress !== p.displayName && ` (${p.emailAddress})`}
                </li>
            )}
          </ul>
        </dd>
      </dl>

      <br/>

      <Button
        href={formUrl}
        target="_blank"
        buttonType={ButtonType.hero}
        iconProps={{ iconName: "ChevronRight" }}
        onClick={click}
      >
        Add to Data Hub
      </Button>

    </div>
  );
}
