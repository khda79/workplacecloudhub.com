> Historical developer reference. The standalone CMDB collection chain is retired.
> Use the product README and PowerBI/README.md for the current prepared consumer.

# Initial CMDB Model

## Entities

- User
- Device
- Group
- License
- Mailbox

## Relationships

- User has assigned license.
- User has mailbox.
- User owns or primarily uses device.
- Device is managed by Intune.
- Group contains user or device.

## Data Quality Signals

- Missing key.
- Duplicate key.
- Stale source.
- Conflicting relationship.
- Low confidence relationship.
