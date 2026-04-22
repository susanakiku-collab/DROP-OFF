// copyLine
// dashboard.js から安全に分離した LINEコピー / LINE送信系
// mobile-safe copy fallback included

function buildCopyResultText() {
  return buildLineResultText();
}

function __copyTextWithTextareaFallback(text) {
  const textarea = document.createElement("textarea");
  textarea.value = String(text || "");
  textarea.setAttribute("readonly", "");
  textarea.setAttribute("aria-hidden", "true");
  textarea.style.position = "fixed";
  textarea.style.top = "0";
  textarea.style.left = "0";
  textarea.style.opacity = "0";
  textarea.style.pointerEvents = "none";
  textarea.style.zIndex = "-1";
  document.body.appendChild(textarea);

  const previousSelection = document.getSelection ? document.getSelection() : null;
  const previousRange =
    previousSelection && previousSelection.rangeCount > 0
      ? previousSelection.getRangeAt(0)
      : null;

  textarea.focus();
  textarea.select();
  textarea.setSelectionRange(0, textarea.value.length);

  let success = false;
  try {
    success = document.execCommand("copy");
  } catch (_) {
    success = false;
  }

  document.body.removeChild(textarea);

  try {
    if (previousSelection) {
      previousSelection.removeAllRanges();
      if (previousRange) previousSelection.addRange(previousRange);
    }
  } catch (_) {}

  return success;
}

async function copyDispatchResult() {
  const text = buildCopyResultText();

  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(text);
      alert("結果をコピーしました");
      return true;
    }
  } catch (error) {
    console.warn("navigator.clipboard failed, fallback to textarea copy:", error);
  }

  const fallbackOk = __copyTextWithTextareaFallback(text);
  if (fallbackOk) {
    alert("結果をコピーしました");
    return true;
  }

  if (navigator.share) {
    try {
      await navigator.share({
        text
      });
      return true;
    } catch (error) {
      console.warn("navigator.share canceled or failed:", error);
    }
  }

  try {
    window.prompt("コピーできない場合は、下の文面を長押ししてコピーしてください", text);
  } catch (_) {}

  alert("コピーに失敗しました。表示された文面を手動でコピーしてください。");
  return false;
}

function sendDispatchResultToLine() {
  const text = buildCopyResultText();
  const url = "https://line.me/R/msg/text/?" + encodeURIComponent(text);
  window.open(url, "_blank");
}
