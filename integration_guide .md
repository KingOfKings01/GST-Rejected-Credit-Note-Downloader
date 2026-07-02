
# 🛡️ Developer Integration Guide: Auth Provider

This guide outlines the step-by-step procedure to integrate custom desktop or client applications with the centralized **Auth Provider** system.

---

## 🔑 Step 1: Provisioning Your Application & Credentials

Before writing code, you must register your application in the administration dashboard.

1. Navigate to the Dashboard URL:
   👉 **[http://13.234.77.157/auth/dashboard](http://13.234.77.157:3005/auth/dashboard)**
2. Click **Create New App** and give it a descriptive name.
3. Once created, copy your secure credentials from the grid:
   - **`App ID`**: Public UUID identifying your application.
   - **`API Key`**: Secret hexadecimal string validating backend authenticity.
4. Ensure you register allowed users (by **email**) under your specific App ID in the **Users** tab.

---

## 📡 Step 2: API Endpoint Implementation

Applications perform user authentication by executing standard JSON POST requests against the Auth Provider APIs.

### 🚀 1. User Authorization Endpoint

Verifies that an email is active, authorized to use your app, and issues a cryptographically signed Token.

- **Method:** `POST`
- **URL:** `http://13.234.77.157/auth/api/authorize`
- **Headers:** `Content-Type: application/json`

#### 📤 Request Body

```json
{
  "app_id": "YOUR-APP-UUID-HERE",
  "api_key": "YOUR-SECRET-API-KEY-HERE",
  "email": "user@example.com"
}
```

#### 📥 Success Response (`200 OK`)

Returns true, registers a `LOGIN` event in audit logs, and issues a secure token.

```json
{
  "authorized": true,
  "message": "Authorized successfully",
  "token": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...",
  "app_name": "Delta Exchange Utility"
}
```

#### ❌ Failure Responses

- **`401 Unauthorized`**: The `api_key` is incorrect.
- **`403 Forbidden`**: The `email` is not registered, or has been set to **Blocked**.
- **`404 Not Found`**: The `app_id` does not match any application in the database.

```json
{
  "authorized": false,
  "message": "Access denied. Your email is blocked."
}
```

---

### 🚪 2. Track Logout Event

Updates the activity logs inside the main dashboard to track exact user durations.

- **Method:** `POST`
- **URL:** `http://13.234.77.157/auth/api/track-logout`
- **Headers:** `Content-Type: application/json`

#### 📤 Request Body

```json
{
  "app_id": "YOUR-APP-UUID-HERE",
  "api_key": "YOUR-SECRET-API-KEY-HERE",
  "email": "user@example.com"
}
```

#### 📥 Success Response (`200 OK`)

```json
{
  "success": true,
  "message": "Logged out tracked"
}
```

---

## 🛠️ Best Practices for Developers

1. [ ] **Client Isolation (JWT Issuer):** The Issued JWT tokens carry the claim `"iss": "auth-provider"`. Always verify this claim on the client backend to avoid credential injection from other projects.
2. [ ] **Error Handling:** If the API returns an `authorized: false` boolean, prompt the user to contact their administrator immediately and prevent application entry.
3. [ ] **Secure Storage:** Never hardcode your `API Key` inside client-side source code that can be easily decompiled. Inject it through secure runtime environment configurations.

**For this project:**

Application ID: 97b89319-db97-4b01-ae9d-3d696f8edc10

Application Key: 7e47552ca953ad5fa46ebcc02648bd64c353cf9d7d945c37
