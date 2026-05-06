import { PROJECTS_ENDPOINT_URL, MOCK_PROJECTS } from "./config";

/* global fetch */

export interface Project {
  id: string;
  name: string;
}

export async function fetchProjects(userEmail: string): Promise<Project[]> {
  if (!PROJECTS_ENDPOINT_URL) {
    // Dev fallback: no real endpoint configured yet.
    return MOCK_PROJECTS;
  }
  const url = `${PROJECTS_ENDPOINT_URL}${PROJECTS_ENDPOINT_URL.includes("?") ? "&" : "?"}userEmail=${encodeURIComponent(
    userEmail
  )}`;
  const res = await fetch(url, { method: "GET", headers: { Accept: "application/json" } });
  if (!res.ok) {
    throw new Error(`Projects API returned HTTP ${res.status} ${res.statusText}`);
  }
  const data = (await res.json()) as Project[];
  if (!Array.isArray(data)) {
    throw new Error("Projects API returned an unexpected payload (expected an array).");
  }
  return data;
}
