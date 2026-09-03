if (window.location.protocol === "file:") {
  window.location.replace("http://127.0.0.1:4173/");
}

const storageKey = "rubric-builder-v1";

const sample = {
  name: "Demo Validation Rubric",
  valueGate: 3,
  topics: [
    { name: "Console fluency & navigation", weight: 15, subtopics: ["Navigate the primary workspace", "Find relevant customer data", "Explain the workflow clearly"] },
    { name: "Setup & onboarding walkthrough", weight: 15, subtopics: ["Set up a new user", "Configure a core setting", "Show a successful first outcome"] },
    { name: "Flagship capability flow", weight: 20, subtopics: ["Start the flagship workflow", "Complete the core action", "Tie it to a customer outcome"] },
    { name: "Core capabilities", weight: 15, subtopics: ["Demonstrate a central capability", "Handle a common use case"] },
    { name: "Visibility & reporting", weight: 10, subtopics: ["Show meaningful reporting", "Connect it to a decision"] },
    { name: "Operational awareness", weight: 10, subtopics: ["Recover from a demo interruption", "Explain operational guardrails"] },
    { name: "Narration & demo hygiene", weight: 15, subtopics: ["Narrate the why, not just the clicks", "Keep a clean presentation state", "Maintain an effective pace"] }
  ]
};

let rubric = load();
const topicsElement = document.querySelector("#topics");
const nameInput = document.querySelector("#rubric-name");
const gateInput = document.querySelector("#value-gate");

function load() {
  try { return JSON.parse(localStorage.getItem(storageKey)) || structuredClone(sample); }
  catch { return structuredClone(sample); }
}
function save() { localStorage.setItem(storageKey, JSON.stringify(rubric)); }
function topicDefaults() { return { name: "New topic", weight: 0, subtopics: ["What must be demonstrated"] }; }

function render() {
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
  addTopicDragEvents(element);
  element.querySelector(".topic-number").textContent = topicIndex + 1;
  const name = element.querySelector(".topic-name"); name.value = topic.name;
  name.addEventListener("input", (event) => { topic.name = event.target.value; save(); });
  const weight = element.querySelector(".topic-weight"); weight.value = topic.weight;
  weight.addEventListener("input", (event) => { topic.weight = Number(event.target.value) || 0; updateWeightStatus(); save(); });
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
  const sideOrb = document.querySelector(".weight-orb");
  const sideMessage = document.querySelector("#side-weight-message");
  sideOrb.classList.toggle("over-limit", overLimit);
  sideMessage.classList.toggle("over-limit", overLimit);
  sideMessage.textContent = overLimit
    ? `Weights exceed 100% by ${weightTotal - 100}%. Reduce a topic weight before exporting.`
    : weightTotal === 100
    ? "Perfect. Every topic is contributing to a complete 100-point rubric."
    : `Your topics total ${weightTotal}%. Adjust them by ${Math.abs(100 - weightTotal)}% to balance the rubric.`;
  document.querySelector("#export-button").disabled = overLimit;
  document.querySelector("#topic-count").textContent = rubric.topics.length;
  document.querySelector("#subtopic-count").textContent = rubric.topics.reduce((total, topic) => total + topic.subtopics.length, 0);
  document.querySelector("#side-value-gate").textContent = `${rubric.valueGate} ${rubric.valueGate === 1 ? "topic" : "topics"}`;
}

nameInput.addEventListener("input", (event) => { rubric.name = event.target.value; save(); });
gateInput.addEventListener("input", (event) => { rubric.valueGate = Math.max(0, Number(event.target.value) || 0); updateWeightStatus(); save(); });
document.querySelector("#add-topic-button").addEventListener("click", () => { rubric.topics.push(topicDefaults()); render(); });
document.querySelector("#reset-button").addEventListener("click", () => { rubric = structuredClone(sample); render(); });
document.querySelector("#export-button").addEventListener("click", () => {
  fetch("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(rubric),
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
