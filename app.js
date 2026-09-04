if (window.location.protocol === "file:") {
  window.location.replace("http://127.0.0.1:4173/");
}

const storageKey = "rubric-builder-v5";

const sample = {
  name: "Demo Validation Rubric",
  valueGate: 3,
  topics: [
    { name: "Console fluency & navigation", weight: 15, subtopics: ["Navigate the primary workspace", "Find relevant customer data", "Explain the workflow clearly"] },
    { name: "Setup & onboarding walkthrough", weight: 15, subtopics: ["Set up a new user", "Configure a core setting", "Show a successful first outcome"] },
    { name: "Flagship capability flow", weight: 20, flagship: true, subtopics: ["Start the flagship workflow", "Complete the core action", "Tie it to a customer outcome"] },
    { name: "Core capabilities", weight: 15, subtopics: ["Demonstrate a central capability", "Handle a common use case"] },
    { name: "Visibility & reporting", weight: 10, subtopics: ["Show meaningful reporting", "Connect it to a decision"] },
    { name: "Operational awareness", weight: 10, subtopics: ["Recover from a demo interruption", "Explain operational guardrails"] },
    { name: "Narration & demo hygiene", weight: 15, subtopics: ["Narrate the why, not just the clicks", "Keep a clean presentation state", "Maintain an effective pace"] }
  ]
};

const samples = {
  demo: sample,
  pitch: {
    name: "Pitch Validation Rubric",
    valueGate: 5,
    topics: [
      { name: "Discovery & qualification", weight: 20, flagship: true, subtopics: ["[Rename] subtopic 1 - what must be shown", "[Rename] subtopic 2 - what must be shown", "[Rename] subtopic 3 - what must be shown", "[Rename] subtopic 4 - what must be shown"] },
      { name: "Features, advantages & benefits", weight: 20, subtopics: ["[Rename] subtopic 1 - what must be shown", "[Rename] subtopic 2 - what must be shown", "[Rename] subtopic 3 - what must be shown", "[Rename] subtopic 4 - what must be shown"] },
      { name: "Common use cases & buyers", weight: 15, subtopics: ["[Rename] subtopic 1 - what must be shown", "[Rename] subtopic 2 - what must be shown", "[Rename] subtopic 3 - what must be shown"] },
      { name: "Competitive positioning", weight: 15, subtopics: ["[Rename] subtopic 1 - what must be shown", "[Rename] subtopic 2 - what must be shown", "[Rename] subtopic 3 - what must be shown"] },
      { name: "Licensing & packaging", weight: 10, subtopics: ["[Rename] subtopic 1 - what must be shown", "[Rename] subtopic 2 - what must be shown", "[Rename] subtopic 3 - what must be shown"] },
      { name: "Delivery & pacing", weight: 20, subtopics: ["Conversation balance - dialogue over monologue; invites and uses interaction", "Pacing & time discipline - tight and energetic; earns its minutes", "Structure & close - clear arc, punchy close, a natural next step"] }
    ]
  },
  poc: {
    name: "POC Validation Rubric",
    valueGate: 3,
    topics: [
      { name: "Environment preparation & staging", weight: 15, subtopics: ["[Rename] Precooked core capability is in place and working before the session", "[Rename] Profiles / policies staged for the POC scenario", "[Rename] Groups or segments prestaged - product-native, not borrowed from AD or an OU", "[Rename] Managed asset enrolled, reporting in, and correctly tagged or classified"] },
      { name: "Integration & dependency verification", weight: 15, subtopics: ["[Rename] The connector or trust relationship is established and authorized on both sides", "[Rename] Shared objects arrive on the consuming device and are visible there", "[Rename] A rule or policy on the consuming side actually uses them"] },
      { name: "[RENAME] FLAGSHIP PROOF UNDER LIVE FAILURE", weight: 20, flagship: true, subtopics: ["[Rename] The continuous check is running and its current state is visible", "[Rename] Induces a real failure on request - not a simulated or narrated one", "[Rename] The state flips and the intended consequence actually lands", "[Rename] Points at the evidence that proves it - the log, the event, the denied session"] },
      { name: "Policy differentiation & segmentation", weight: 15, subtopics: ["[Rename] Different groups receive genuinely different policy - segmentation proven, not described", "[Rename] Distinguishes two commonly confused capabilities by showing them behave differently", "[Rename] Pushes a change and confirms it landed on the target"] },
      { name: "Troubleshooting", weight: 20, subtopics: ["[Rename] Diagnoses a broken enrollment, telemetry, or connectivity dependency", "[Rename] Diagnoses a broken access or enforcement path - trust, mismatch, scope, or ordering", "[Rename] Uses the right evidence sources in the right order rather than guessing"] },
      { name: "Scoping & engagement judgment (walkthrough accepted)", weight: 15, subtopics: ["[Rename] Walks the configuration for an integration that need not be built in-session", "[Rename] Explains the migration path off an incumbent agent, and the rollback story", "[Rename] Names when to recommend the Best Practice Service (BPS) instead of proceeding solo"] }
    ]
  }
};

const trackMeta = {
  demo: { label: "Demo", exportReady: true, gateLabel: "Value-tied topics required", note: "" },
  pitch: { label: "Pitch", exportReady: true, gateLabel: "Tailoring gate minimum", note: "" },
  poc: { label: "POC", exportReady: true, gateLabel: "Validation gate minimum", note: "" },
};

let state = load();
let activeTrack = state.activeTrack;
let rubric = state.rubrics[activeTrack];
const topicsElement = document.querySelector("#topics");
const nameInput = document.querySelector("#rubric-name");
const gateInput = document.querySelector("#value-gate");
const gateLabel = document.querySelector("label[for='value-gate'] span") || document.querySelector(".setup-card label:nth-child(2) span");
const trackNote = document.querySelector("#track-note");
const trackTabs = Array.from(document.querySelectorAll(".track-tab"));

function load() {
  const fallback = { activeTrack: "demo", rubrics: structuredClone(samples) };
  try {
    const stored = JSON.parse(localStorage.getItem(storageKey));
    if (stored?.rubrics) return normalizeState(stored);
    const v4 = JSON.parse(localStorage.getItem("rubric-builder-v4"));
    if (v4?.rubrics) return normalizeState(v4);
    const v3 = JSON.parse(localStorage.getItem("rubric-builder-v3"));
    if (v3?.rubrics) {
      fallback.activeTrack = v3.activeTrack || "demo";
      fallback.rubrics.demo = v3.rubrics.demo || fallback.rubrics.demo;
      fallback.rubrics.pitch = v3.rubrics.pitch || fallback.rubrics.pitch;
      return normalizeState(fallback);
    }
    const v2 = JSON.parse(localStorage.getItem("rubric-builder-v2"));
    if (v2?.rubrics) {
      fallback.activeTrack = v2.activeTrack === "poc" ? "demo" : (v2.activeTrack || "demo");
      fallback.rubrics.demo = v2.rubrics.demo || fallback.rubrics.demo;
      return normalizeState(fallback);
    }
    const legacy = JSON.parse(localStorage.getItem("rubric-builder-v1"));
    if (legacy?.topics) fallback.rubrics.demo = legacy;
    return normalizeState(fallback);
  } catch {
    return normalizeState(fallback);
  }
}

function normalizeState(rawState) {
  const next = {
    activeTrack: trackMeta[rawState.activeTrack] ? rawState.activeTrack : "demo",
    rubrics: structuredClone(samples),
  };
  Object.keys(samples).forEach((track) => {
    if (rawState.rubrics?.[track]) next.rubrics[track] = rawState.rubrics[track];
    next.rubrics[track].topics = (next.rubrics[track].topics || []).map((topic, index) => ({
      ...topic,
      flagship: Boolean(topic.flagship ?? samples[track].topics[index]?.flagship),
    }));
  });
  return next;
}

function save() {
  state.activeTrack = activeTrack;
  state.rubrics[activeTrack] = rubric;
  localStorage.setItem(storageKey, JSON.stringify(state));
}
function topicDefaults() { return { name: "New topic", weight: 0, flagship: false, subtopics: ["What must be demonstrated"] }; }

function render() {
  const meta = trackMeta[activeTrack];
  trackTabs.forEach((tab) => tab.classList.toggle("active", tab.dataset.track === activeTrack));
  gateLabel.textContent = meta.gateLabel;
  trackNote.hidden = !meta.note;
  trackNote.textContent = meta.note;
  nameInput.value = rubric.name;
  gateInput.value = rubric.valueGate;
  topicsElement.innerHTML = "";
  rubric.topics.forEach((topic, topicIndex) => topicsElement.append(makeTopic(topic, topicIndex)));
  updateWeightStatus();
  save();
}

function makeTopic(topic, topicIndex) {
  const element = document.querySelector("#topic-template").content.firstElementChild.cloneNode(true);
  element.dataset.topicIndex = topicIndex;
  element.classList.toggle("flagship", Boolean(topic.flagship));
  addTopicDragEvents(element);
  element.querySelector(".topic-number").textContent = topicIndex + 1;
  const name = element.querySelector(".topic-name"); name.value = topic.name;
  name.addEventListener("input", (event) => { topic.name = event.target.value; save(); });
  const weight = element.querySelector(".topic-weight"); weight.value = topic.weight;
  weight.addEventListener("input", (event) => { topic.weight = Number(event.target.value) || 0; updateWeightStatus(); save(); });
  const flagship = element.querySelector(".topic-flagship");
  flagship.checked = Boolean(topic.flagship);
  flagship.addEventListener("change", (event) => {
    rubric.topics.forEach((candidate) => { candidate.flagship = false; });
    topic.flagship = event.target.checked;
    render();
  });
  element.querySelector(".delete-topic").addEventListener("click", () => { rubric.topics.splice(topicIndex, 1); render(); });
  const list = element.querySelector(".subtopic-list");
  topic.subtopics.forEach((subtopic, subtopicIndex) => list.append(makeSubtopic(topic, subtopic, subtopicIndex)));
  element.querySelector(".add-subtopic").addEventListener("click", () => { topic.subtopics.push({ name: "New evidence item", level: "" }); render(); });
  return element;
}

function addTopicDragEvents(element) {
  const handle = element.querySelector(".drag-handle");
  handle.addEventListener("pointerdown", () => { element.draggable = true; });
  handle.addEventListener("pointerup", () => { element.draggable = false; });
  element.addEventListener("dragstart", (event) => {
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", element.dataset.topicIndex);
    requestAnimationFrame(() => element.classList.add("dragging"));
  });
  element.addEventListener("dragend", () => {
    element.draggable = false;
    document.querySelectorAll(".topic-card").forEach((card) => card.classList.remove("dragging", "drop-target"));
  });
  element.addEventListener("dragover", (event) => {
    event.preventDefault();
    const source = document.querySelector(".topic-card.dragging");
    if (source && source !== element) element.classList.add("drop-target");
  });
  element.addEventListener("dragleave", () => element.classList.remove("drop-target"));
  element.addEventListener("drop", (event) => {
    event.preventDefault();
    const sourceIndex = Number(event.dataTransfer.getData("text/plain"));
    const targetIndex = Number(element.dataset.topicIndex);
    if (Number.isInteger(sourceIndex) && sourceIndex !== targetIndex) {
      const [movedTopic] = rubric.topics.splice(sourceIndex, 1);
      rubric.topics.splice(targetIndex, 0, movedTopic);
      render();
    }
  });
}

function makeSubtopic(topic, rawSubtopic, subtopicIndex) {
  const subtopic = typeof rawSubtopic === "string" ? { name: rawSubtopic } : rawSubtopic;
  topic.subtopics[subtopicIndex] = subtopic;
  const element = document.querySelector("#subtopic-template").content.firstElementChild.cloneNode(true);
  const name = element.querySelector(".subtopic-name"); name.value = subtopic.name;
  name.addEventListener("input", (event) => { subtopic.name = event.target.value; save(); });
  element.querySelector(".delete-subtopic").addEventListener("click", () => { topic.subtopics.splice(subtopicIndex, 1); render(); });
  return element;
}

function updateWeightStatus() {
  const weightTotal = rubric.topics.reduce((total, topic) => total + (Number(topic.weight) || 0), 0);
  const weightStatus = document.querySelector("#weight-status");
  weightStatus.textContent = weightTotal === 100 ? "Weights are balanced at 100%." : `Weights total ${weightTotal}%. Add or remove ${Math.abs(100 - weightTotal)}%.`;
  weightStatus.classList.toggle("invalid", weightTotal !== 100);
  document.querySelector("#side-weight-total").textContent = weightTotal;
  const overLimit = weightTotal > 100;
  const exportReady = trackMeta[activeTrack].exportReady;
  const sideOrb = document.querySelector(".weight-orb");
  const sideMessage = document.querySelector("#side-weight-message");
  sideOrb.classList.toggle("over-limit", overLimit);
  sideMessage.classList.toggle("over-limit", overLimit);
  sideMessage.textContent = overLimit
    ? `Weights exceed 100% by ${weightTotal - 100}%. Reduce a topic weight before exporting.`
    : weightTotal === 100
    ? "Perfect. Every topic is contributing to a complete 100-point rubric."
    : `Your topics total ${weightTotal}%. Adjust them by ${Math.abs(100 - weightTotal)}% to balance the rubric.`;
  document.querySelector("#export-button").disabled = overLimit || !exportReady;
  document.querySelector("#topic-count").textContent = rubric.topics.length;
  document.querySelector("#subtopic-count").textContent = rubric.topics.reduce((total, topic) => total + topic.subtopics.length, 0);
  document.querySelector("#side-value-gate").textContent = `${rubric.valueGate} ${rubric.valueGate === 1 ? "topic" : "topics"}`;
}

nameInput.addEventListener("input", (event) => { rubric.name = event.target.value; save(); });
gateInput.addEventListener("input", (event) => { rubric.valueGate = Math.max(0, Number(event.target.value) || 0); updateWeightStatus(); save(); });
document.querySelector("#add-topic-button").addEventListener("click", () => { rubric.topics.push(topicDefaults()); render(); });
document.querySelector("#reset-button").addEventListener("click", () => { rubric = structuredClone(samples[activeTrack]); render(); });
trackTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    save();
    activeTrack = tab.dataset.track;
    rubric = state.rubrics[activeTrack];
    render();
  });
});
document.querySelector("#export-button").addEventListener("click", () => {
  if (!trackMeta[activeTrack].exportReady) {
    alert(`${trackMeta[activeTrack].label} export will be enabled after its layout is added.`);
    return;
  }
  fetch("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ...rubric, track: activeTrack }),
  })
    .then(async (response) => {
      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.error || "Excel export is unavailable.");
      }
      return response.blob();
    })
    .then((blob) => {
      const link = Object.assign(document.createElement("a"), {
        href: URL.createObjectURL(blob),
        download: `${rubric.name || "rubric"}.xlsx`,
      });
      link.click();
      URL.revokeObjectURL(link.href);
    })
    .catch((error) => alert(error.message));
});

render();
