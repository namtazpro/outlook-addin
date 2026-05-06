import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  Button,
  Dropdown,
  Field,
  MessageBar,
  MessageBarBody,
  Option,
  Spinner,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { CloudArrowUp24Regular } from "@fluentui/react-icons";
import { buildEmailPayload } from "../emailPayload";
import { UPLOAD_ENDPOINT_URL } from "../config";
import { fetchProjects, Project } from "../projectsApi";

/* global Office, fetch */

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    alignItems: "stretch",
    gap: "12px",
    padding: "20px",
  },
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    textAlign: "center",
  },
  buttonRow: {
    display: "flex",
    justifyContent: "center",
  },
  status: {
    width: "100%",
  },
});

type UploadStatus =
  | { kind: "idle" }
  | { kind: "uploading" }
  | { kind: "success"; bytes: number; project: string }
  | { kind: "error"; message: string };

type ProjectsState =
  | { kind: "loading" }
  | { kind: "ready"; projects: Project[] }
  | { kind: "error"; message: string };

const SendEmailToBlob: React.FC = () => {
  const styles = useStyles();
  const [uploadStatus, setUploadStatus] = useState<UploadStatus>({ kind: "idle" });
  const [projectsState, setProjectsState] = useState<ProjectsState>({ kind: "loading" });
  const [selectedProjectId, setSelectedProjectId] = useState<string | null>(null);

  const userEmail = useMemo<string>(
    () => Office.context.mailbox?.userProfile?.emailAddress ?? "",
    []
  );

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const projects = await fetchProjects(userEmail);
        if (cancelled) return;
        setProjectsState({ kind: "ready", projects });
      } catch (err) {
        if (cancelled) return;
        setProjectsState({
          kind: "error",
          message: err instanceof Error ? err.message : String(err),
        });
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [userEmail]);

  const selectedProject: Project | null =
    projectsState.kind === "ready"
      ? projectsState.projects.find((p) => p.id === selectedProjectId) ?? null
      : null;

  const handleClick = async () => {
    if (!selectedProject) return;
    setUploadStatus({ kind: "uploading" });
    try {
      const payload = await buildEmailPayload({ id: selectedProject.id, name: selectedProject.name });
      const body = JSON.stringify(payload);
      const response = await fetch(UPLOAD_ENDPOINT_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body,
      });
      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(`HTTP ${response.status} ${response.statusText}${text ? `: ${text}` : ""}`);
      }
      setUploadStatus({ kind: "success", bytes: body.length, project: selectedProject.name });
    } catch (err) {
      setUploadStatus({
        kind: "error",
        message: err instanceof Error ? err.message : String(err),
      });
    }
  };

  return (
    <div className={styles.root}>
      <Field className={styles.instructions}>
        Capture this email and upload it as JSON to Azure Blob Storage.
      </Field>

      <Field label="Project">
        {projectsState.kind === "loading" && <Spinner size="tiny" label="Loading projects\u2026" />}
        {projectsState.kind === "error" && (
          <MessageBar intent="error">
            <MessageBarBody>Failed to load projects: {projectsState.message}</MessageBarBody>
          </MessageBar>
        )}
        {projectsState.kind === "ready" && (
          <Dropdown
            placeholder="Select a project"
            value={selectedProject?.name ?? ""}
            selectedOptions={selectedProjectId ? [selectedProjectId] : []}
            onOptionSelect={(_e, data) => setSelectedProjectId(data.optionValue ?? null)}
          >
            {projectsState.projects.map((p) => (
              <Option key={p.id} value={p.id} text={p.name}>
                {p.name}
              </Option>
            ))}
          </Dropdown>
        )}
      </Field>

      <div className={styles.buttonRow}>
        <Button
          appearance="primary"
          size="large"
          icon={<CloudArrowUp24Regular />}
          disabled={uploadStatus.kind === "uploading" || !selectedProject}
          onClick={handleClick}
        >
          {uploadStatus.kind === "uploading" ? "Uploading\u2026" : "Send email to Blob Storage"}
        </Button>
      </div>

      {uploadStatus.kind === "success" && (
        <MessageBar intent="success" className={styles.status}>
          <MessageBarBody>
            Uploaded {uploadStatus.bytes.toLocaleString()} bytes for {uploadStatus.project}.
          </MessageBarBody>
        </MessageBar>
      )}
      {uploadStatus.kind === "error" && (
        <MessageBar intent="error" className={styles.status}>
          <MessageBarBody>Upload failed: {uploadStatus.message}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
};

export default SendEmailToBlob;
