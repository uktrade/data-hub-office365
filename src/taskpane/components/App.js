import * as React from 'react'
import { useState } from 'react'
import {
  Button,
  ButtonType,
  Link,
  MessageBar,
  MessageBarType,
  Spinner,
} from 'office-ui-fabric-react'
/* global Button, Header, HeroList, HeroListItem, Progress */


const CALENDAR_ITEM_CLASSES = [
  'IPM.Schedule.Meeting.Request',
]

export default function App({ title, isOfficeInitialized }) {
  const [ created, setCreated ] = useState(false)
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

  return (
    <div style={{ margin: '5%' }}>
      <h3>Interaction details</h3>

      <dl>
        <dt>Subject</dt><dd>{item.subject}</dd>
        <dt>Location</dt><dd>{item.location}</dd>
        <dt>Date</dt><dd>{item.start.format()}</dd>
        <dt>Participants</dt>
        <dd>
          <ul>
            {item.to.map(p =>
              <li key={p.emailAddress}>
                {p.displayName}{p.emailAddress !== p.displayName && ` (${p.emailAddress})`}
              </li>)}
          </ul>
        </dd>
      </dl>

      <br/>

      {!created && (
        <Button
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          onClick={click}
        >
          Add to Data Hub
        </Button>
      )}

      {created && (
        <MessageBar>
          Interaction was created.
          <Link href="http://example.com" target="_blank">
            View it on Data Hub.
          </Link>
        </MessageBar>
      )}

    </div>
  );
}
