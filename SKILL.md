# M365 Email Draft Skill

Create a draft email in Microsoft 365.

## Usage
`uv run main.py --subject {{subject}} --body {{body}} --to {{to}} --cc {{cc}} --bcc {{bcc}}`

## Parameters
- `to`: One or more recipient email addresses (space-separated).
- `cc`: carbon copy email addresses (space-separated) (optional).
- `bcc`: blind carbon copy recipient email addresses (space-separated) (optional).
- `subject`: The subject of the email.
- `body`: The content of the email (supports HTML).
