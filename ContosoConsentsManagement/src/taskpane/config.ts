// Endpoints are injected at build time by webpack (DefinePlugin) from environment
// variables / a local `.env` file. See `.env.example` and `webpack.config.js`.
//
// `process.env.*` is replaced with string literals at build time; the `declare`
// below just keeps TypeScript happy.

declare const process: { env: { UPLOAD_ENDPOINT_URL?: string; PROJECTS_ENDPOINT_URL?: string } };

// Full Azure Function URL (including ?code=<function key>) that receives the email payload.
export const UPLOAD_ENDPOINT_URL: string = process.env.UPLOAD_ENDPOINT_URL ?? "";

// Endpoint of the API that returns the list of projects available to a user.
// Called as: GET `${PROJECTS_ENDPOINT_URL}?userEmail=...`
// Leave empty to use MOCK_PROJECTS below.
export const PROJECTS_ENDPOINT_URL: string = process.env.PROJECTS_ENDPOINT_URL ?? "";

// Mock projects returned when PROJECTS_ENDPOINT_URL is empty (development).
export const MOCK_PROJECTS = [
  { id: "alpha", name: "Project Alpha" },
  { id: "beta", name: "Project Beta" },
  { id: "gamma", name: "Project Gamma" },
];
