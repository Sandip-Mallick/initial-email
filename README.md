# Initial Email Outlook Add-in

An Outlook Add-in that allows you to extract the body text of a selected email and save it as a JSON file.

## Features

- Extract email body content from the selected email in Outlook
- Save email data (subject, sender, date, body) as a JSON file
- Works with Outlook Desktop (Windows/Mac) and Outlook Web

## Prerequisites

- [Node.js](https://nodejs.org) (LTS version recommended)
- Microsoft Outlook (Desktop or Web)
- OpenSSL (for certificate generation)

## Setup for Development

1. Clone or download this repository
2. Navigate to the project directory
3. Install the dependencies:
   ```
   npm install
   ```
4. Generate self-signed certificates for HTTPS development:
   ```
   # Create certs directory
   mkdir -p certs
   cd certs

   # Generate server key and certificate
   openssl req -newkey rsa:2048 -nodes -keyout server.key -x509 -days 365 -out server.crt -subj "/CN=localhost"

   # Create certificate serial number file
   echo "01" > ca.srl

   # Generate CA key
   openssl genrsa -out ca.key 2048

   # Generate CA certificate
   openssl req -new -x509 -key ca.key -out ca.crt -days 365 -subj "/CN=localhost"

   # Return to project root
   cd ..
   ```
5. Start the development server:
   ```
   npm start
   ```

## Sideloading the Add-in in Outlook

### Outlook Desktop (Windows)
1. Open Outlook
2. Click on the gear icon (Settings) and select "Manage Add-ins"
3. Click "My add-ins" in the left navigation
4. Click the "+" icon and select "Add from file..."
5. Browse to the `manifest.xml` file in the project root and select it
6. Follow the prompts to install the add-in

### Outlook Desktop (Mac)
1. Open Outlook
2. Click "Tools" > "Add-ins"
3. Click "My add-ins" in the left navigation
4. Click the "+" icon and select "Add from file..."
5. Browse to the `manifest.xml` file in the project root and select it
6. Follow the prompts to install the add-in

### Outlook Web
1. Go to [Outlook Web](https://outlook.office.com)
2. Click on the gear icon (Settings) and select "Manage Add-ins"
3. Click "My add-ins" in the left navigation
4. Click the "+" icon and select "Add from file..."
5. Browse to the `manifest.xml` file in the project root and select it
6. Follow the prompts to install the add-in

## Usage

1. Open Outlook and select an email
2. Click the "Initial Email" button in the ribbon
3. In the taskpane that appears, click "Save Email as JSON"
4. The email content will be downloaded as a JSON file to your local system

## Troubleshooting

### Certificate Issues
If you encounter certificate-related errors:
1. Make sure you've generated all the required certificate files in the `certs` directory
2. Verify that the certificates are valid and not expired
3. Check that the certificates are properly referenced in your development environment

## License

ISC
