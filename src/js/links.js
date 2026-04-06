// ═══════════════════════════════════════════════════════════════
//  DEPENDENCY LINK SYSTEM
// ═══════════════════════════════════════════════════════════════

/** Toggle "Link Mode" – shows anchor dots on every bar */
function toggleLinkMode() {
  linkMode = !linkMode;
  linkSource = null; // reset pending selection
  const btn = document.getElementById("link-mode-btn");
  const banner = document.getElementById("link-mode-banner");
  if (linkMode) {
    btn.style.background = "#6366f1";
    btn.style.color = "white";
    if (banner) {
      banner.style.display = "block";
      banner.textContent =
        "🔗 Link Mode ON – click a start/end dot to begin";
    }
  } else {
    btn.style.background = "";
    btn.style.color = "";
    if (banner) banner.style.display = "none";
  }
  render();
}

/** Toggle visibility of all link SVG lines */
function toggleShowLinks() {
  showLinks = !showLinks;
  const btn = document.getElementById("show-links-btn");
  const linksLabel = document.getElementById("links-label");
  if (linksLabel) linksLabel.textContent = showLinks ? "Links: On" : "Links: Off";
  else if (btn) btn.textContent = showLinks ? "👁 Links: On" : "👁 Links: Off";
  drawLinks();
}

/**
 * Two-click anchor selection handler.
 * First click → stores linkSource.
 * Second click → creates the link (if different item or opposite anchor).
 */
function clickAnchor(itemId, anchor) {
  if (!linkSource) {
    // First click – store source
    linkSource = { itemId, anchor };
    const banner = document.getElementById("link-mode-banner");
    if (banner)
      banner.textContent =
        "🔗 Source selected (" + anchor + ") – now click the target dot";
    // Refresh to show active dot highlight
    render();
    return;
  }

  // Second click
  if (linkSource.itemId === itemId && linkSource.anchor === anchor) {
    // Clicked same dot → deselect
    linkSource = null;
    render();
    return;
  }

  // Prevent duplicate links
  const exists = links.find(
    (l) =>
      l.fromId === linkSource.itemId &&
      l.fromAnchor === linkSource.anchor &&
      l.toId === itemId &&
      l.toAnchor === anchor,
  );
  if (!exists) {
    pushUndo();
    links.push({
      id: nextLinkId++,
      fromId: linkSource.itemId,
      fromAnchor: linkSource.anchor,
      toId: itemId,
      toAnchor: anchor,
    });
  }

  linkSource = null;
  const banner = document.getElementById("link-mode-banner");
  if (banner)
    banner.textContent =
      "🔗 Link Mode ON – click a start/end dot to begin";
  render();
}

/** Delete a specific link by id */
function deleteLink(linkId) {
  pushUndo();
  links = links.filter((l) => l.id !== linkId);
  render();
}

/**
 * Draw SVG link arrows.
 * Uses position:fixed so getBoundingClientRect() viewport coords map directly
 * to SVG coords — no scroll math required.
 * Redraws on plannerEl horizontal scroll and on window scroll.
 */
function drawLinks() {
  const old = document.getElementById("link-svg-overlay");
  if (old) old.remove();
  if (!showLinks || links.length === 0) return;

  // CSS zoom on body scales content but BCR returns viewport (screen) coords.
  // The fixed SVG is inside body and therefore also zoomed, so its coordinate
  // system differs from viewport coords. Dividing BCR by zoom converts screen
  // coords into the SVG's CSS-pixel space.
  const zoom = parseFloat(getComputedStyle(document.body).zoom) || 1;

  // Get the planner bounds early so we can build a clipPath
  const plannerEl = gridEl.closest
    ? gridEl.closest(".planner-container")
    : gridEl.parentElement;
  const rawClip = plannerEl ? plannerEl.getBoundingClientRect() : null;
  const clipRect = rawClip ? {
    left: rawClip.left / zoom,
    top: rawClip.top / zoom,
    width: rawClip.width / zoom,
    height: rawClip.height / zoom,
  } : null;

  // Fixed SVG covers the entire viewport — BCR coords (adjusted for zoom) work directly
  const svg = document.createElementNS(
    "http://www.w3.org/2000/svg",
    "svg",
  );
  svg.id = "link-svg-overlay";
  svg.style.cssText = [
    "position:fixed",
    "top:0",
    "left:0",
    "width:100vw",
    "height:100vh",
    "pointer-events:none",
    "z-index:12",
    "overflow:visible",
  ].join(";");
  document.body.appendChild(svg);

  // Arrow marker + clip path scoped to the planner container
  const defs = document.createElementNS(
    "http://www.w3.org/2000/svg",
    "defs",
  );
  const marker = document.createElementNS(
    "http://www.w3.org/2000/svg",
    "marker",
  );
  marker.setAttribute("id", "link-arrow");
  marker.setAttribute("markerWidth", "8");
  marker.setAttribute("markerHeight", "8");
  marker.setAttribute("refX", "6");
  marker.setAttribute("refY", "3");
  marker.setAttribute("orient", "auto");
  const arrowPath = document.createElementNS(
    "http://www.w3.org/2000/svg",
    "path",
  );
  arrowPath.setAttribute("d", "M0,0 L0,6 L8,3 z");
  arrowPath.setAttribute("fill", "#6366f1");
  marker.appendChild(arrowPath);
  defs.appendChild(marker);

  if (clipRect) {
    const clipPath = document.createElementNS(
      "http://www.w3.org/2000/svg",
      "clipPath",
    );
    clipPath.setAttribute("id", "link-clip-area");
    const clipRectEl = document.createElementNS(
      "http://www.w3.org/2000/svg",
      "rect",
    );
    clipRectEl.setAttribute("x", clipRect.left);
    clipRectEl.setAttribute("y", clipRect.top);
    clipRectEl.setAttribute("width", clipRect.width);
    clipRectEl.setAttribute("height", clipRect.height);
    clipPath.appendChild(clipRectEl);
    defs.appendChild(clipPath);
  }

  svg.appendChild(defs);

  // All link elements go into a clipped group
  const linesGroup = document.createElementNS(
    "http://www.w3.org/2000/svg",
    "g",
  );
  if (clipRect) linesGroup.setAttribute("clip-path", "url(#link-clip-area)");
  svg.appendChild(linesGroup);

  links.forEach((link) => {
    const fromBlock = document.getElementById("block-" + link.fromId);
    const toBlock = document.getElementById("block-" + link.toId);
    if (!fromBlock || !toBlock) return;

    const fr = fromBlock.getBoundingClientRect();
    const tr = toBlock.getBoundingClientRect();

    // Convert viewport (screen) coords to the SVG's CSS-pixel space
    const sx = (link.fromAnchor === "start" ? fr.left : fr.right) / zoom;
    const sy = (fr.top + fr.height / 2) / zoom;
    const tx = (link.toAnchor === "start" ? tr.left : tr.right) / zoom;
    const ty = (tr.top + tr.height / 2) / zoom;

    const pull = Math.min(Math.abs(tx - sx) * 0.5, 120);
    const csx = sx + (link.fromAnchor === "end" ? pull : -pull);
    const ctx2 = tx + (link.toAnchor === "start" ? -pull : pull);

    const midX = (sx + tx) / 2;
    const midY = (sy + ty) / 2;

    const g = document.createElementNS("http://www.w3.org/2000/svg", "g");
    g.setAttribute("class", "link-line-group");

    // Wide transparent hit area
    const hitPath = document.createElementNS(
      "http://www.w3.org/2000/svg",
      "path",
    );
    hitPath.setAttribute(
      "d",
      `M${sx},${sy} C${csx},${sy} ${ctx2},${ty} ${tx},${ty}`,
    );
    hitPath.setAttribute("stroke", "transparent");
    hitPath.setAttribute("stroke-width", "12");
    hitPath.setAttribute("fill", "none");
    g.appendChild(hitPath);

    // Visible dashed bezier
    const visPath = document.createElementNS(
      "http://www.w3.org/2000/svg",
      "path",
    );
    visPath.setAttribute(
      "d",
      `M${sx},${sy} C${csx},${sy} ${ctx2},${ty} ${tx},${ty}`,
    );
    visPath.setAttribute("stroke", "#6366f1");
    visPath.setAttribute("stroke-width", "2.5");
    visPath.setAttribute("fill", "none");
    visPath.setAttribute("stroke-dasharray", "7,4");
    visPath.setAttribute("marker-end", "url(#link-arrow)");
    g.appendChild(visPath);

    // Red ✕ delete button at midpoint
    const delCircle = document.createElementNS(
      "http://www.w3.org/2000/svg",
      "circle",
    );
    delCircle.setAttribute("cx", midX);
    delCircle.setAttribute("cy", midY);
    delCircle.setAttribute("r", "9");
    delCircle.setAttribute("fill", "#ef4444");
    delCircle.setAttribute("pointer-events", "all");
    delCircle.style.cursor = "pointer";
    const lId = link.id;
    delCircle.addEventListener("click", (e) => {
      e.stopPropagation();
      deleteLink(lId);
    });

    const delText = document.createElementNS(
      "http://www.w3.org/2000/svg",
      "text",
    );
    delText.setAttribute("x", midX);
    delText.setAttribute("y", midY + 4);
    delText.setAttribute("text-anchor", "middle");
    delText.setAttribute("font-size", "11");
    delText.setAttribute("fill", "white");
    delText.setAttribute("font-weight", "bold");
    delText.setAttribute("pointer-events", "none");
    delText.textContent = "✕";

    g.appendChild(delCircle);
    g.appendChild(delText);
    linesGroup.appendChild(g);
  });

  // Register scroll handlers so wires follow bars when the user scrolls
  if (plannerEl) plannerEl.onscroll = drawLinks;
  window.onscroll = drawLinks;
}

/**
 * Cascade date changes through dependency links.
 *
 * Link semantics (fromAnchor -> toAnchor):
 *   end   -> start  = Finish-to-Start  (FS) — most common
 *   start -> start  = Start-to-Start   (SS)
 *   end   -> end    = Finish-to-Finish (FF)
 *   start -> end    = Start-to-Finish  (SF)
 *
 * @param {number} movedId    – ID of the item whose dates changed
 * @param {number} deltaStart – how much the START of movedId shifted (weeks)
 * @param {number} deltaEnd   – how much the END   of movedId shifted (weeks)
 * @param {Set}    visited    – cycle guard (pass new Set() on first call)
 */
function propagateLinks(movedId, deltaStart, deltaEnd, visited) {
  if (!visited) visited = new Set();
  if (visited.has(movedId)) return;
  visited.add(movedId);

  links.forEach((link) => {
    // Only propagate FORWARD (from predecessor to successor)
    if (link.fromId !== movedId) return;

    const target = items.find((i) => i.id === link.toId);
    if (!target) return;
    if (visited.has(target.id)) return;

    // Skip targets that were already moved as part of the drag cluster
    if (draggingItems && draggingItems.includes(target.id)) return;

    // Pick the relevant delta based on which anchor of the predecessor changed
    const delta = link.fromAnchor === "start" ? deltaStart : deltaEnd;
    if (delta === 0) return;

    const oldTargetStart = target.startWeek;
    const oldTargetEnd = oldTargetStart + (target.duration || 0);

    // Apply the shift to the successor's startWeek (all link types move successor in time)
    const newStart = parseFloat((target.startWeek + delta).toFixed(1));
    if (newStart >= 1 && newStart <= getTotalWeekCount()) {
      target.startWeek = newStart;
    }

    // Compute the actual deltas applied to the successor for recursive propagation
    const appliedDeltaStart = parseFloat(
      (target.startWeek - oldTargetStart).toFixed(1),
    );
    const newTargetEnd = target.startWeek + (target.duration || 0);
    const appliedDeltaEnd = parseFloat(
      (newTargetEnd - oldTargetEnd).toFixed(1),
    );

    // Cascade to successors of the successor
    propagateLinks(
      target.id,
      appliedDeltaStart,
      appliedDeltaEnd,
      visited,
    );
  });
}
