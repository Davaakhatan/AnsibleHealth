# Authenticate and import required libraries
from google.colab import auth

auth.authenticate_user()  # Prompts to authenticate via Google Account

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Build the Google Docs service object
docs_service = build('docs', 'v1')

# Define the markdown meeting notes (this can also be read from a file)
markdown_text = """
# Product Team Sync - May 15, 2023

## Attendees
- Sarah Chen (Product Lead)
- Mike Johnson (Engineering)
- Anna Smith (Design)
- David Park (QA)

## Agenda

### 1. Sprint Review
* Completed Features
  * User authentication flow
  * Dashboard redesign
  * Performance optimization
    * Reduced load time by 40%
    * Implemented caching solution
* Pending Items
  * Mobile responsive fixes
  * Beta testing feedback integration

### 2. Current Challenges
* Resource constraints in QA team
* Third-party API integration delays
* User feedback on new UI
  * Navigation confusion
  * Color contrast issues

### 3. Next Sprint Planning
* Priority Features
  * Payment gateway integration
  * User profile enhancement
  * Analytics dashboard
* Technical Debt
  * Code refactoring
  * Documentation updates

## Action Items
- [ ] @sarah: Finalize Q3 roadmap by Friday
- [ ] @mike: Schedule technical review for payment integration
- [ ] @anna: Share updated design system documentation
- [ ] @david: Prepare QA resource allocation proposal

## Next Steps
* Schedule individual team reviews
* Update sprint board
* Share meeting summary with stakeholders

## Notes
* Next sync scheduled for May 22, 2023
* Platform demo for stakeholders on May 25
* Remember to update JIRA tickets

---
Meeting recorded by: Sarah Chen
Duration: 45 minutes
"""


def parse_markdown_to_requests(md_text):
    """
    Parses a markdown string and creates a list of Google Docs API requests
    to insert text with appropriate formatting.
    """
    requests = []
    # Track the current insertion index in the document
    current_index = 1  # Google Docs index starts at 1

    # Split markdown into individual lines
    for line in md_text.splitlines():
        # Skip empty lines
        if not line.strip():
            continue

        # Check for main title (Heading 1)
        if line.startswith("# "):
            text = line[2:].strip() + "\n"
            requests.append({
                'insertText': {
                    'location': {'index': current_index},
                    'text': text
                }
            })
            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': current_index,
                        'endIndex': current_index + len(text)
                    },
                    'paragraphStyle': {
                        'namedStyleType': 'HEADING_1'
                    },
                    'fields': 'namedStyleType'
                }
            })
            current_index += len(text)

        # Section header (Heading 2)
        elif line.startswith("## "):
            text = line[3:].strip() + "\n"
            requests.append({
                'insertText': {
                    'location': {'index': current_index},
                    'text': text
                }
            })
            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': current_index,
                        'endIndex': current_index + len(text)
                    },
                    'paragraphStyle': {
                        'namedStyleType': 'HEADING_2'
                    },
                    'fields': 'namedStyleType'
                }
            })
            current_index += len(text)

        # Sub-section header (Heading 3)
        elif line.startswith("### "):
            text = line[4:].strip() + "\n"
            requests.append({
                'insertText': {
                    'location': {'index': current_index},
                    'text': text
                }
            })
            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': current_index,
                        'endIndex': current_index + len(text)
                    },
                    'paragraphStyle': {
                        'namedStyleType': 'HEADING_3'
                    },
                    'fields': 'namedStyleType'
                }
            })
            current_index += len(text)

        # Checkbox item (convert "- [ ]" into a checkbox bullet)
        elif line.strip().startswith("- [ ]"):
            # Remove the markdown checkbox marker and trim
            text = line.strip().replace("- [ ]", "").strip() + "\n"
            requests.append({
                'insertText': {
                    'location': {'index': current_index},
                    'text': text
                }
            })
            # Create a bullet list with a checkbox preset
            requests.append({
                'createParagraphBullets': {
                    'range': {
                        'startIndex': current_index,
                        'endIndex': current_index + len(text)
                    },
                    'bulletPreset': 'BULLET_CHECKBOX'
                }
            })
            current_index += len(text)

        # Regular bullet list item (using "-" or "*")
        elif line.lstrip().startswith("-") or line.lstrip().startswith("*"):
            # Determine indentation level (for nested bullets)
            indent_level = (len(line) - len(line.lstrip(' '))) // 2
            # Remove the bullet marker and trim
            text = line.lstrip(" -*").strip() + "\n"
            requests.append({
                'insertText': {
                    'location': {'index': current_index},
                    'text': text
                }
            })
            requests.append({
                'createParagraphBullets': {
                    'range': {
                        'startIndex': current_index,
                        'endIndex': current_index + len(text)
                    },
                    'bulletPreset': 'BULLET_DISC'
                }
            })
            # Adjust indentation if needed
            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': current_index,
                        'endIndex': current_index + len(text)
                    },
                    'paragraphStyle': {
                        'indentStart': {
                            'magnitude': indent_level * 18,
                            'unit': 'PT'
                        }
                    },
                    'fields': 'indentStart'
                }
            })
            current_index += len(text)

        # Footer separator (assumes '---' as separator)
        elif line.startswith("---"):
            # Insert a newline (separator) without additional styling
            text = "\n"
            requests.append({
                'insertText': {
                    'location': {'index': current_index},
                    'text': text
                }
            })
            current_index += len(text)

        # Footer information (Meeting recorded by, Duration)
        elif line.startswith("Meeting recorded by:") or line.startswith("Duration:"):
            text = line.strip() + "\n"
            requests.append({
                'insertText': {
                    'location': {'index': current_index},
                    'text': text
                }
            })
            # Apply italic style with a gray color to footer info
            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': current_index,
                        'endIndex': current_index + len(text)
                    },
                    'textStyle': {
                        'italic': True,
                        'foregroundColor': {
                            'color': {
                                'rgbColor': {'red': 0.5, 'green': 0.5, 'blue': 0.5}
                            }
                        }
                    },
                    'fields': 'italic,foregroundColor'
                }
            })
            current_index += len(text)

        # All other lines – normal paragraphs (also handle @mentions)
        else:
            text = line.strip() + "\n"
            requests.append({
                'insertText': {
                    'location': {'index': current_index},
                    'text': text
                }
            })
            # Check for @mentions within the text and apply bold style with a custom color
            mention_idx = text.find('@')
            while mention_idx != -1:
                # Find the end of the mention (up to the next space or end-of-line)
                end_idx = mention_idx
                while end_idx < len(text) and text[end_idx] not in [' ', '\n']:
                    end_idx += 1
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': current_index + mention_idx,
                            'endIndex': current_index + end_idx
                        },
                        'textStyle': {
                            'bold': True,
                            'foregroundColor': {
                                'color': {
                                    'rgbColor': {'red': 0.2, 'green': 0.4, 'blue': 0.8}
                                }
                            }
                        },
                        'fields': 'bold,foregroundColor'
                    }
                })
                # Look for next occurrence in the same line
                mention_idx = text.find('@', end_idx)
            current_index += len(text)

    return requests


# Create a new Google Doc
try:
    doc_title = "Product Team Sync - May 15, 2023"
    create_response = docs_service.documents().create(body={'title': doc_title}).execute()
    document_id = create_response.get('documentId')
    print("Document created successfully!")
except HttpError as error:
    print(f"An error occurred while creating the document: {error}")
    raise error

# Parse the markdown text into a list of Google Docs API requests
requests = parse_markdown_to_requests(markdown_text)

# Update the document with the parsed formatting requests
try:
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
    print("Document updated successfully!")
except HttpError as error:
    print(f"An error occurred while updating the document: {error}")
    raise error

# Output the document URL
doc_url = f"https://docs.google.com/document/d/{document_id}"
print("Access your document here:", doc_url)
