// ═══════════════════════════════════════════════════════════════
//  DRAG, RESIZE, PAN, ROW REORDER
// ═══════════════════════════════════════════════════════════════

// ── Planner pan (click-drag to scroll) ─────────────────────────

function startPan(e) {
  if (e.button !== 0) return;
  if (e.target.closest(
    'button, input, a, select, textarea, ' +
    '.bar, .range-bar, .resize-handle, .col-resizer, ' +
    '.link-anchor, .settings-panel, .comment-block, ' +
    '.milestone, .milestone-diamond, .milestone-wrapper, ' +
    '.name-cell, .comment-cell, .drag-handle, .comment-popup'
  )) return;
  const container = e.currentTarget;
  isPanning = true;
  panStartX = e.clientX;
  panStartY = e.clientY;
  panScrollLeft = container.scrollLeft;
  panScrollTop = container.scrollTop;
  container.classList.add("is-panning");
  e.preventDefault();
  document.addEventListener("mousemove", onPan);
  document.addEventListener("mouseup", stopPan);
}

function onPan(e) {
  if (!isPanning) return;
  const container = document.querySelector(".planner-container");
  container.scrollLeft = panScrollLeft - (e.clientX - panStartX);
  container.scrollTop = panScrollTop - (e.clientY - panStartY);
}

function stopPan() {
  if (!isPanning) return;
  isPanning = false;
  const container = document.querySelector(".planner-container");
  container.classList.remove("is-panning");
  document.removeEventListener("mousemove", onPan);
  document.removeEventListener("mouseup", stopPan);
}

// ── Bar Resize ─────────────────────────────────────────────────

function startResize(e, id, type) {
  if (resizingId !== null) return; // guard: don't double-register
  const item = items.find((i) => i.id === id);
  if (!item || item.isLocked === true) return;
  e.preventDefault();
  e.stopPropagation();
  resizingId = id;
  resizeType = type;
  resizeStartX = e.clientX;
  initialStartWeek = item.startWeek;
  initialDuration = item.duration;

  _resizeEl = document.getElementById("block-" + id);
  _resizeInitialBarW = _resizeEl ? _resizeEl.offsetWidth : 0;
  _resizeInitialBarML = _resizeEl ? (parseFloat(_resizeEl.style.marginLeft) || 0) : 0;
  _resizeNewStart = item.startWeek;
  _resizeNewDuration = item.duration;

  // Promote to compositor layer — transform-only updates during resize
  if (_resizeEl) {
    _resizeEl.style.willChange = "transform";
    _resizeEl.style.transformOrigin = "left center";
  }

  document.addEventListener("mousemove", onResize);
  document.addEventListener("mouseup", stopResize);
  document.body.style.cursor = "ew-resize";
}

function onResize(e) {
  if (!resizingId || !_resizeEl) return;

  const dx = e.clientX - resizeStartX;
  let cellWidth = 28;
  if (zoomMode === "days") cellWidth = 20;
  if (zoomMode === "months") cellWidth = 80;

  const snapCells = Math.round(dx / cellWidth);
  const pxDelta = snapCells * cellWidth;

  let weeksMoved = 0;
  if (zoomMode === "weeks") weeksMoved = snapCells;
  else if (zoomMode === "days") weeksMoved = snapCells * 0.2;
  else if (zoomMode === "months") {
    if (resizeType === "left") {
      const targetStart = getAbsWeekFromMonthPosition(getMonthPositionFromAbsWeek(initialStartWeek) + snapCells);
      weeksMoved = parseFloat((targetStart - initialStartWeek).toFixed(1));
    } else {
      const initialEndWeek = initialStartWeek + initialDuration;
      const targetEnd = getAbsWeekFromMonthPosition(getMonthPositionFromAbsWeek(initialEndWeek) + snapCells);
      weeksMoved = parseFloat((targetEnd - initialEndWeek).toFixed(1));
    }
  }

  if (resizeType === "left") {
    let newStart = parseFloat((initialStartWeek + weeksMoved).toFixed(1));
    let newDuration = parseFloat((initialDuration - weeksMoved).toFixed(1));

    if (newStart < 1) { newDuration -= 1 - newStart; newStart = 1; }
    if (newDuration < 0.2) { newStart = initialStartWeek + initialDuration - 0.2; newDuration = 0.2; }

    _resizeNewStart = newStart;
    _resizeNewDuration = newDuration;

    // compositor-only: translateX shifts the left edge, scaleX adjusts width from that edge
    const clampedDelta = Math.min(pxDelta, _resizeInitialBarW - cellWidth);
    const factor = Math.max(0.01, (_resizeInitialBarW - clampedDelta) / _resizeInitialBarW);
    if (!_interactionRafId) {
      _interactionRafId = requestAnimationFrame(() => {
        _interactionRafId = null;
        _resizeEl.style.transform = "translateX(" + clampedDelta + "px) scaleX(" + factor + ")";
      });
    }
  } else if (resizeType === "right") {
    let newDuration = parseFloat((initialDuration + weeksMoved).toFixed(1));
    if (newDuration < 0.2) newDuration = 0.2;
    if (initialStartWeek + newDuration - 1 > getTotalWeekCount()) {
      newDuration = getTotalWeekCount() - initialStartWeek + 1;
    }

    _resizeNewDuration = newDuration;

    // compositor-only: scaleX stretches/shrinks from the left edge
    const newW = Math.max(cellWidth, _resizeInitialBarW + pxDelta);
    const factor = newW / _resizeInitialBarW;
    if (!_interactionRafId) {
      _interactionRafId = requestAnimationFrame(() => {
        _interactionRafId = null;
        _resizeEl.style.transform = "scaleX(" + factor + ")";
      });
    }
  }
}

function stopResize() {
  if (_interactionRafId) {
    cancelAnimationFrame(_interactionRafId);
    _interactionRafId = null;
  }

  if (resizingId !== null) {
    // Clear all visual overrides — render() will set correct values
    if (_resizeEl) {
      _resizeEl.style.transform = "";
      _resizeEl.style.transformOrigin = "";
      _resizeEl.style.willChange = "";
    }

    const item = items.find((i) => i.id === resizingId);
    if (item) {
      const deltaStart = parseFloat((_resizeNewStart - initialStartWeek).toFixed(1));
      const oldEnd = initialStartWeek + initialDuration;
      const newEnd = _resizeNewStart + _resizeNewDuration;
      const deltaEnd = parseFloat((newEnd - oldEnd).toFixed(1));
      if (deltaStart !== 0 || deltaEnd !== 0) {
        pushUndo();
        item.startWeek = _resizeNewStart;
        item.duration = _resizeNewDuration;
        propagateLinks(resizingId, deltaStart, deltaEnd, new Set());
        render();
      }
    }
  }

  _resizeEl = null;
  _resizeNewStart = 0;
  _resizeNewDuration = 0;
  resizingId = null;
  resizeType = null;
  document.removeEventListener("mousemove", onResize);
  document.removeEventListener("mouseup", stopResize);
  document.body.style.cursor = "default";
}

// ── Bar Drag ───────────────────────────────────────────────────

function handleBlockClick(e, itemId) {
  if (e.ctrlKey) {
    e.preventDefault();
    e.stopPropagation();
    const item = items.find((i) => i.id === itemId);
    if (item) {
      if (!item.blockComment) {
        const gridEl = document.getElementById("grid");
        const gridRect = gridEl.getBoundingClientRect();
        const scrollParent = gridEl.closest(".planner-container") || gridEl.parentElement;
        const scrollX = scrollParent ? scrollParent.scrollLeft : 0;
        const scrollY = scrollParent ? scrollParent.scrollTop : 0;
        item.blockComment = {
          text: "",
          x: e.clientX - gridRect.left + scrollX,
          y: e.clientY - gridRect.top + scrollY,
          width: 200,
          height: 150,
          isOpen: true,
        };
      } else {
        item.blockComment.isOpen = !item.blockComment.isOpen;
      }
      render();
    }
    return true;
  }
  return false;
}

function startDrag(e, id) {
  if (handleBlockClick(e, id)) return;
  if (draggingId !== null) return; // guard: don't double-register
  const item = items.find((i) => i.id === id);
  if (!item || item.isLocked === true) return;

  e.preventDefault();
  draggingId = id;
  dragStartX = e.clientX;
  dragLatestClientX = e.clientX;
  const dragContainer = document.querySelector(".planner-container");
  dragStartScrollLeft = dragContainer ? dragContainer.scrollLeft : 0;

  draggingItems = [id];
  initialStartWeeks = {};
  _dragStartViewportBounds = new Map();
  _dragPreviewSnapPx = 0;

  // Build O(1) lookup structures once for the entire drag session
  _dragItemsMap = new Map(items.map(i => [i.id, i]));
  const _dragChildMap = new Map();
  items.forEach(i => {
    if (i.parentId) {
      if (!_dragChildMap.has(i.parentId)) _dragChildMap.set(i.parentId, []);
      _dragChildMap.get(i.parentId).push(i);
    }
  });

  const _kids = _dragChildMap.get(id) || [];
  if (_kids.length > 0) {
    let _min = Infinity;
    const _scan = (pid) =>
      (_dragChildMap.get(pid) || []).forEach((k) => {
        if (k.startWeek < _min) _min = k.startWeek;
        _scan(k.id);
      });
    _scan(id);
    if (_min !== Infinity) item.startWeek = _min;
  }
  initialStartWeeks[id] = item.startWeek;

  function addDescendants(parentId) {
    const children = _dragChildMap.get(parentId) || [];
    children.forEach((child) => {
      if (!draggingItems.includes(child.id)) {
        draggingItems.push(child.id);
        initialStartWeeks[child.id] = child.startWeek;
        addDescendants(child.id);
      }
    });
  }
  addDescendants(id);

  // Promote dragged bars to their own compositor layer for smooth transforms
  draggingItems.forEach((itemId) => {
    const el = document.getElementById("block-" + itemId);
    if (el) {
      if (typeof el.getBoundingClientRect === "function") {
        const rect = el.getBoundingClientRect();
        _dragStartViewportBounds.set(itemId, { left: rect.left, right: rect.right });
      }
      el.style.willChange = "transform";
      el.classList.add("is-dragging");
    }
  });

  document.addEventListener("mousemove", onDrag);
  document.addEventListener("mouseup", stopDrag);
  document.body.style.cursor = "grabbing";
  _startDragAutoScroll();
}

function _getDraggedBlockViewportBounds() {
  if (!draggingItems || draggingItems.length === 0) return null;
  let minLeft = Infinity;
  let maxRight = -Infinity;
  draggingItems.forEach((itemId) => {
    const startRect = _dragStartViewportBounds.get(itemId);
    if (!startRect) return;
    const left = startRect.left + _dragPreviewSnapPx;
    const right = startRect.right + _dragPreviewSnapPx;
    if (left < minLeft) minLeft = left;
    if (right > maxRight) maxRight = right;
  });
  if (minLeft === Infinity || maxRight === -Infinity) return null;
  return { left: minLeft, right: maxRight };
}

function _getDragAutoScrollDelta(clientX) {
  const container = document.querySelector(".planner-container");
  const grid = document.getElementById("grid");
  if (!container || !grid) return 0;

  const containerRect = container.getBoundingClientRect();
  const stickyNameCell = grid.querySelector(".header-row-4 .name-cell") || grid.querySelector(".name-cell");
  const firstTlHeader = grid.querySelector(".tl-header");
  const leftThreshold = stickyNameCell
    ? stickyNameCell.getBoundingClientRect().right
    : firstTlHeader
      ? firstTlHeader.getBoundingClientRect().left
      : containerRect.left;
  const rightThreshold = containerRect.right - 24;
  const maxStep = 28;
  const draggedBounds = _getDraggedBlockViewportBounds();
  const leftProbe = draggedBounds ? draggedBounds.left : clientX;

  if (leftProbe <= leftThreshold) {
    const strength = Math.min(1, (leftThreshold - leftProbe + 8) / 80);
    return -Math.max(8, Math.round(maxStep * strength));
  }
  if (clientX > rightThreshold) {
    const strength = Math.min(1, (clientX - rightThreshold) / 80);
    return Math.max(8, Math.round(maxStep * strength));
  }
  return 0;
}

function _updateDragPosition(clientX) {
  if (!draggingId) return;
  const container = document.querySelector(".planner-container");
  const scrollLeft = container ? container.scrollLeft : 0;
  const dx = clientX - dragStartX + (scrollLeft - dragStartScrollLeft);
  let cellWidth = 28;
  if (zoomMode === "days") cellWidth = 20;
  if (zoomMode === "months") cellWidth = 80;

  const snapCells = Math.round(dx / cellWidth);

  let weeksMoved = 0;
  if (zoomMode === "weeks") {
    weeksMoved = snapCells;
  } else if (zoomMode === "days") {
    weeksMoved = snapCells * 0.2;
  } else if (zoomMode === "months") {
    const baseStartWeek = initialStartWeeks[draggingId] || 1;
    const targetStart = getAbsWeekFromMonthPosition(getMonthPositionFromAbsWeek(baseStartWeek) + snapCells);
    weeksMoved = parseFloat((targetStart - baseStartWeek).toFixed(1));
  }

  let minAllowedMoved = -Infinity;
  let maxAllowedMoved = Infinity;

  draggingItems.forEach((itemId) => {
    const initialStart = initialStartWeeks[itemId];
    if (initialStart >= 1) {
      const minMove = 1 - initialStart;
      if (minMove > minAllowedMoved) minAllowedMoved = minMove;
    }
    if (initialStart <= getTotalWeekCount()) {
      const maxMove = getTotalWeekCount() - initialStart;
      if (maxMove < maxAllowedMoved) maxAllowedMoved = maxMove;
    }
  });
  if (maxAllowedMoved < minAllowedMoved) weeksMoved = 0;
  if (weeksMoved < minAllowedMoved) weeksMoved = minAllowedMoved;
  if (weeksMoved > maxAllowedMoved) weeksMoved = maxAllowedMoved;

  finalWeeksMoved = weeksMoved;

  // Visual-only move: CSS transform on each bar — no DOM rebuild, no render().
  // item.startWeek is committed in stopDrag once the user releases.
  // Pixel delta = weeksMoved converted back to cells × cellWidth, preserving clamping.
  let snapPx = 0;
  if (zoomMode === "weeks") snapPx = weeksMoved * cellWidth;
  else if (zoomMode === "days") snapPx = weeksMoved * 5 * cellWidth;   // 1 week = 5 day-cells × 20px
  else if (zoomMode === "months") snapPx = snapCells * cellWidth;
  _dragPreviewSnapPx = snapPx;
  if (!_interactionRafId) {
    _interactionRafId = requestAnimationFrame(() => {
      _interactionRafId = null;
      draggingItems.forEach((itemId) => {
        const el = document.getElementById("block-" + itemId);
        if (el) el.style.transform = "translateX(" + snapPx + "px)";
      });
    });
  }
}

function onDrag(e) {
  if (!draggingId) return;
  dragLatestClientX = e.clientX;
  _updateDragPosition(e.clientX);
}

function _startDragAutoScroll() {
  if (_dragAutoScrollRafId) cancelAnimationFrame(_dragAutoScrollRafId);
  const tick = () => {
    if (!draggingId) {
      _dragAutoScrollRafId = null;
      return;
    }
    const container = document.querySelector(".planner-container");
    if (container) {
      const delta = _getDragAutoScrollDelta(dragLatestClientX);
      if (delta !== 0) {
        const maxScrollLeft = Math.max(0, container.scrollWidth - container.clientWidth);
        const nextScrollLeft = Math.max(0, Math.min(maxScrollLeft, container.scrollLeft + delta));
        if (nextScrollLeft !== container.scrollLeft) {
          container.scrollLeft = nextScrollLeft;
          _updateDragPosition(dragLatestClientX);
        }
      }
    }
    _dragAutoScrollRafId = requestAnimationFrame(tick);
  };
  _dragAutoScrollRafId = requestAnimationFrame(tick);
}

function stopDrag() {
  if (_interactionRafId) {
    cancelAnimationFrame(_interactionRafId);
    _interactionRafId = null;
  }
  if (_dragAutoScrollRafId) {
    cancelAnimationFrame(_dragAutoScrollRafId);
    _dragAutoScrollRafId = null;
  }

  // Clear visual transforms and will-change on all dragged bars
  draggingItems.forEach((itemId) => {
    const el = document.getElementById("block-" + itemId);
    if (el) {
      el.style.transform = "";
      el.style.willChange = "";
      el.classList.remove("is-dragging");
    }
  });

  if (draggingId !== null && finalWeeksMoved !== 0) {
    pushUndo();
    // Commit positions — item.startWeek was NOT modified during drag
    draggingItems.forEach((itemId) => {
      const item = _dragItemsMap.get(itemId);
      if (item) item.startWeek = parseFloat((initialStartWeeks[itemId] + finalWeeksMoved).toFixed(1));
    });
    const visited = new Set();
    draggingItems.forEach((itemId) => {
      propagateLinks(itemId, finalWeeksMoved, finalWeeksMoved, visited);
    });
    render();
  }

  finalWeeksMoved = 0;
  draggingId = null;
  dragLatestClientX = 0;
  dragStartScrollLeft = 0;
  _dragPreviewSnapPx = 0;
  _dragStartViewportBounds = new Map();
  document.removeEventListener("mousemove", onDrag);
  document.removeEventListener("mouseup", stopDrag);
  document.body.style.cursor = "default";
}

// ── Column Resize ──────────────────────────────────────────────

function startColResize(e, type) {
  e.preventDefault();
  e.stopPropagation();
  colResizing = type;
  colResizeStartX = e.clientX;
  colResizeStartWidth = type === "comment" ? commentWidth : nameWidthBase;

  document.addEventListener("mousemove", onColResize);
  document.addEventListener("mouseup", stopColResize);
  document.body.style.cursor = "col-resize";
}

function onColResize(e) {
  if (!colResizing) return;
  const dx = e.clientX - colResizeStartX;
  const newWidth = Math.max(50, colResizeStartWidth + dx);

  if (colResizing === "comment") {
    commentWidth = newWidth;
  } else if (colResizing === "name") {
    nameWidthBase = newWidth;
  }

  const actualCommentWidth = showComments ? commentWidth + "px" : "0px";
  const actualNameWidth = nameWidthBase + (showSettings ? 280 : 0) + "px";

  gridEl.style.setProperty("--comment-width", actualCommentWidth);
  gridEl.style.setProperty("--name-width", actualNameWidth);
}

function stopColResize() {
  colResizing = null;
  document.removeEventListener("mousemove", onColResize);
  document.removeEventListener("mouseup", stopColResize);
  document.body.style.cursor = "default";
  render();
}

// ── Block Comment Drag/Resize ──────────────────────────────────

function startCommentDrag(e, itemId) {
  e.preventDefault();
  e.stopPropagation();
  const item = items.find((i) => i.id === itemId);
  if (!item) return;

  if (!item.blockComment) {
    const gridEl = document.getElementById("grid");
    const gridRect = gridEl.getBoundingClientRect();
    const scrollParent = gridEl.closest(".planner-container") || gridEl.parentElement;
    const scrollX = scrollParent ? scrollParent.scrollLeft : 0;
    const scrollY = scrollParent ? scrollParent.scrollTop : 0;
    item.blockComment = {
      text: "",
      x: e.clientX - gridRect.left + scrollX,
      y: e.clientY - gridRect.top + scrollY,
      width: 200,
      height: 150,
      isOpen: true,
    };
  }

  draggingCommentId = itemId;
  commentDragStartX = e.clientX;
  commentDragStartY = e.clientY;
  commentInitialX = item.blockComment.x || 0;
  commentInitialY = item.blockComment.y || 0;

  document.addEventListener("mousemove", onCommentDrag);
  document.addEventListener("mouseup", stopCommentInteraction);
}

function onCommentDrag(e) {
  if (draggingCommentId === null) return;
  const item = items.find((i) => i.id === draggingCommentId);
  if (!item || !item.blockComment) return;

  const dx = e.clientX - commentDragStartX;
  const dy = e.clientY - commentDragStartY;

  item.blockComment.x = commentInitialX + dx;
  item.blockComment.y = commentInitialY + dy;

  const el = document.getElementById(`comment-popup-${draggingCommentId}`);
  if (el) {
    el.style.left = item.blockComment.x + "px";
    el.style.top = item.blockComment.y + "px";
    updatePopupPointer(item, el);
  }
}

function startCommentResize(e, itemId) {
  e.preventDefault();
  e.stopPropagation();
  const item = items.find((i) => i.id === itemId);
  if (!item || !item.blockComment) return;

  resizingCommentId = itemId;
  commentDragStartX = e.clientX;
  commentDragStartY = e.clientY;
  commentInitialW = item.blockComment.width || 200;
  commentInitialH = item.blockComment.height || 150;

  document.addEventListener("mousemove", onCommentResize);
  document.addEventListener("mouseup", stopCommentInteraction);
}

function onCommentResize(e) {
  if (resizingCommentId === null) return;
  const item = items.find((i) => i.id === resizingCommentId);
  if (!item || !item.blockComment) return;

  const dx = e.clientX - commentDragStartX;
  const dy = e.clientY - commentDragStartY;

  item.blockComment.width = Math.max(150, commentInitialW + dx);
  item.blockComment.height = Math.max(100, commentInitialH + dy);

  const el = document.getElementById(`comment-popup-${resizingCommentId}`);
  if (el) {
    el.style.width = item.blockComment.width + "px";
    el.style.height = item.blockComment.height + "px";
    updatePopupPointer(item, el);
  }
}

function stopCommentInteraction() {
  draggingCommentId = null;
  resizingCommentId = null;
  document.removeEventListener("mousemove", onCommentDrag);
  document.removeEventListener("mousemove", onCommentResize);
  document.removeEventListener("mouseup", stopCommentInteraction);
}

function updateBlockCommentText(itemId, text) {
  const item = items.find((i) => i.id === itemId);
  if (item && item.blockComment) {
    item.blockComment.text = text;
  }
}

function closeComment(itemId) {
  const item = items.find((i) => i.id === itemId);
  if (item && item.blockComment) {
    item.blockComment.isOpen = false;
    render();
  }
}

function deleteComment(itemId) {
  if (confirm("Delete this comment?")) {
    const item = items.find((i) => i.id === itemId);
    if (item) {
      delete item.blockComment;
      render();
    }
  }
}

function getCommentTailSVG() {
  let svg = document.getElementById("comment-tail-svg-container");
  if (!svg) {
    svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.id = "comment-tail-svg-container";
    svg.classList.add("comment-tail-svg");
    document.getElementById("grid").appendChild(svg);
  }
  return svg;
}

function _gridOffset(el) {
  let x = 0, y = 0;
  const grid = document.getElementById("grid");
  let cur = el;
  while (cur && cur !== grid) {
    x += cur.offsetLeft;
    y += cur.offsetTop;
    cur = cur.offsetParent;
  }
  return { x, y };
}

function updatePopupPointer(item, el) {
  if (!item.blockComment) return;
  const blockEl = document.getElementById(`block-${item.id}`);
  if (!blockEl) return;

  const svg = getCommentTailSVG();
  let path = document.getElementById(`comment-tail-${item.id}`);
  if (!path) {
    path = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
    path.id = `comment-tail-${item.id}`;
    path.setAttribute("fill", "white");
    path.setAttribute("stroke", "#cbd5e1");
    path.setAttribute("stroke-width", "1");
    svg.appendChild(path);
  }

  const bOff = _gridOffset(blockEl);
  const bMidX = bOff.x + blockEl.offsetWidth / 2;
  const bMidY = bOff.y + blockEl.offsetHeight / 2;

  const pX = el.offsetLeft;
  const pY = el.offsetTop;
  const pW = el.offsetWidth;
  const pH = el.offsetHeight;
  const pMidX = pX + pW / 2;
  const pMidY = pY + pH / 2;

  const dx = bMidX - pMidX;
  const dy = bMidY - pMidY;

  let startX1, startY1, startX2, startY2;
  const baseWidth = 20;

  if (Math.abs(dx) > Math.abs(dy)) {
    if (dx > 0) {
      startX1 = pX + pW; startY1 = pMidY - baseWidth / 2;
      startX2 = pX + pW; startY2 = pMidY + baseWidth / 2;
    } else {
      startX1 = pX; startY1 = pMidY - baseWidth / 2;
      startX2 = pX; startY2 = pMidY + baseWidth / 2;
    }
  } else {
    if (dy > 0) {
      startX1 = pMidX - baseWidth / 2; startY1 = pY + pH;
      startX2 = pMidX + baseWidth / 2; startY2 = pY + pH;
    } else {
      startX1 = pMidX - baseWidth / 2; startY1 = pY;
      startX2 = pMidX + baseWidth / 2; startY2 = pY;
    }
  }

  path.setAttribute("points",
    `${startX1},${startY1} ${startX2},${startY2} ${bMidX},${bMidY}`);
}

// ── Row Drag & Drop ────────────────────────────────────────────

function startRowDrag(e, id) {
  draggedRowId = id;
  e.dataTransfer.effectAllowed = "move";
  e.dataTransfer.setData("text/plain", id);
}

function rowDragOver(e) {
  e.preventDefault();
  const nameCell = e.currentTarget;
  if (e.altKey) {
    nameCell.classList.add("drag-nest-over");
    nameCell.classList.remove("drag-unnest-over");
  } else if (e.ctrlKey || e.metaKey) {
    nameCell.classList.add("drag-unnest-over");
    nameCell.classList.remove("drag-nest-over");
  } else {
    nameCell.classList.remove("drag-nest-over");
    nameCell.classList.remove("drag-unnest-over");
  }
  e.dataTransfer.dropEffect = "move";
}

function rowDragLeave(e) {
  e.currentTarget.classList.remove("drag-nest-over");
  e.currentTarget.classList.remove("drag-unnest-over");
}

function rowDragOverBottom(e) {
  e.preventDefault();
  e.currentTarget.style.backgroundColor = "#eef2ff";
  e.currentTarget.style.color = "#4f46e5";
  e.dataTransfer.dropEffect = "move";
}

function rowDragLeaveBottom(e) {
  e.currentTarget.style.backgroundColor = "";
  e.currentTarget.style.color = "#cbd5e1";
}

function rowDrop(e, targetId) {
  e.preventDefault();
  e.currentTarget.classList.remove("drag-nest-over");
  e.currentTarget.classList.remove("drag-unnest-over");
  e.currentTarget.style.backgroundColor = "";
  e.currentTarget.style.color = "#cbd5e1";

  if (draggedRowId === null || draggedRowId === targetId) {
    draggedRowId = null;
    return;
  }

  const draggedIdx = items.findIndex((i) => i.id === draggedRowId);
  if (draggedIdx === -1) { draggedRowId = null; return; }

  const draggedItem = items[draggedIdx];

  // Helper function to collect all descendant IDs of the dragged item
  function getSubtreeIds(rootId) {
    const ids = new Set([rootId]);
    let changed = true;
    while (changed) {
      changed = false;
      items.forEach((i) => {
        if (i.parentId && ids.has(i.parentId) && !ids.has(i.id)) {
          ids.add(i.id);
          changed = true;
        }
      });
    }
    return ids;
  }

  // BOTTOM DROP
  if (targetId === 'bottom') {
    pushUndo();
    try {
      if (draggedItem.parentId) delete draggedItem.parentId;
      const subtreeIds = getSubtreeIds(draggedItem.id);
      const subtreeItems = items.filter((i) => subtreeIds.has(i.id));
      items = items.filter((i) => !subtreeIds.has(i.id));
      items.push(...subtreeItems);
      render();
    } catch (err) {
      console.error("Error in bottom drop:", err);
    }
    draggedRowId = null;
    return;
  }

  const targetIdx = items.findIndex((i) => i.id === targetId);
  if (targetIdx === -1) { draggedRowId = null; return; }

  // ALT-DROP: Re-parent
  if (e.altKey) {
    pushUndo();
    if (getSubtreeIds(draggedItem.id).has(targetId)) {
      draggedRowId = null;
      return;
    }
    draggedItem.parentId = targetId;
    const subtreeIds = getSubtreeIds(draggedItem.id);
    const subtreeItems = items.filter((i) => subtreeIds.has(i.id));
    items = items.filter((i) => !subtreeIds.has(i.id));
    const newTargetIdx = items.findIndex((i) => i.id === targetId);
    items.splice(newTargetIdx + 1, 0, ...subtreeItems);
    render();
    draggedRowId = null;
    return;
  }

  // CTRL-DROP: Un-nest
  if (e.ctrlKey || e.metaKey) {
    pushUndo();
    draggedItem.parentId = null;
    const subtreeIds = getSubtreeIds(draggedItem.id);
    const subtreeItems = items.filter((i) => subtreeIds.has(i.id));
    items = items.filter((i) => !subtreeIds.has(i.id));
    const newTargetIdx = items.findIndex((i) => i.id === targetId);
    items.splice(newTargetIdx + 1, 0, ...subtreeItems);
    render();
    draggedRowId = null;
    return;
  }

  // NORMAL DROP: Reorder within same parent
  let tId = targetId;
  let tIdx = items.findIndex((i) => i.id === tId);
  let targetItem = items[tIdx];

  while (targetItem && draggedItem.parentId !== targetItem.parentId) {
    if (!targetItem.parentId) { targetItem = null; break; }
    tId = targetItem.parentId;
    tIdx = items.findIndex((i) => i.id === tId);
    targetItem = items[tIdx];
  }

  if (targetItem && draggedItem.parentId === targetItem.parentId) {
    pushUndo();
    items.splice(draggedIdx, 1);
    const newTIdx = items.findIndex((i) => i.id === tId);
    if (draggedIdx < targetIdx) {
      items.splice(newTIdx + 1, 0, draggedItem);
    } else {
      items.splice(newTIdx, 0, draggedItem);
    }
    render();
  }

  draggedRowId = null;
}

function getCommentLinks(text) {
  if (!text) return [];
  const linkRegex = /(https?:\/\/[^\s]+|file:\/\/[^\s]+|[a-zA-Z]:\\[^\s]+|\\\\[^\s]+)/g;
  return [...new Set(text.match(linkRegex) || [])];
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getCommentLinksHtml(text, compact) {
  const links = getCommentLinks(text);
  if (links.length === 0) return "";
  const cls = compact ? "comment-link-list compact" : "comment-link-list";
  return (
    '<div class="' + cls + '">' +
    links
      .slice(0, compact ? 1 : 6)
      .map((link) => {
        const safeLink = escapeHtml(link);
        return '<a class="comment-detected-link" href="#" onclick="openCommentLink(\'' + safeLink + '\'); return false;" title="' + safeLink + '">' + safeLink + "</a>";
      })
      .join("") +
    "</div>"
  );
}

function openCommentLink(link) {
  if (!link) return;
  window.open(link, "_blank");
}

// ── Open URLs and Local Folder paths via click inside comments ──
function handleCommentLinkClick(e) {
  const text = e.target.value;
  if (!text) return;
  if (e.ctrlKey) {
    const links = getCommentLinks(text);
    if (links.length === 1) {
      e.preventDefault();
      e.stopPropagation();
      openCommentLink(links[0]);
    }
  }
}
