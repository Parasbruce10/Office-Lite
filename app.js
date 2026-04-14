// app.js (with React)

// --- Helper for File Downloads ---
const downloadFile = (content, fileName, mimeType) => {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
};

// --- Editor Components ---

// Simplified Word Editor
const WordEditor = ({ onClose }) => {
    const fileInputRef = React.useRef(null);
    const editorRef = React.useRef(null);
    const savedSelectionRef = React.useRef(null);

    // ✅ Selection Save/Restore System
    const saveSelection = () => {
        const sel = window.getSelection();
        if (sel.rangeCount > 0) {
            savedSelectionRef.current = sel.getRangeAt(0).cloneRange();
        }
    };

    const restoreSelection = () => {
        const editor = editorRef.current;
        if (editor) editor.focus();
        const sel = window.getSelection();
        if (savedSelectionRef.current) {
            try {
                sel.removeAllRanges();
                sel.addRange(savedSelectionRef.current);
            } catch (err) { }
        }
    };

    const execute = (cmd, val = null) => {
        restoreSelection();
        document.execCommand(cmd, false, val);
        saveSelection();
    };

    const keepFocus = (e) => {
        e.preventDefault();
    };

    // ========================================================
    // ✅ IMAGE RESIZE & MOVE SYSTEM
    // ========================================================
    React.useEffect(() => {
        const editor = editorRef.current;
        if (!editor) return;

        let activeImg = null;
        let overlay = null;
        let action = null;
        let startX, startY, startW, startH, startML, startMT;

        const removeOverlay = () => {
            if (overlay && overlay.parentNode) {
                overlay.parentNode.removeChild(overlay);
            }
            overlay = null;
        };

        const clearImgSelection = () => {
            if (activeImg) {
                activeImg.style.outline = 'none';
                activeImg.style.cursor = '';
            }
            activeImg = null;
            removeOverlay();
        };

        // ✅ FIX: getBoundingClientRect use karo — ye hamesha sahi position deta hai
        const updateOverlayPosition = () => {
            if (!activeImg || !overlay) return;
            const editorRect = editor.getBoundingClientRect();
            const imgRect = activeImg.getBoundingClientRect();
            overlay.style.top = (imgRect.top - editorRect.top) + 'px';
            overlay.style.left = (imgRect.left - editorRect.left) + 'px';
            overlay.style.width = imgRect.width + 'px';
            overlay.style.height = imgRect.height + 'px';
        };

        const createOverlay = (img) => {
            clearImgSelection();
            activeImg = img;
            img.style.outline = '3px solid #2a5fbd';

            overlay = document.createElement('div');
            overlay.setAttribute('data-overlay', 'true');
            overlay.style.cssText = 'position:absolute;z-index:100;box-sizing:border-box;pointer-events:none;';

            // 4 corner resize handles
            const corners = [
                { name: 'nw', css: 'top:-6px;left:-6px;cursor:nw-resize;' },
                { name: 'ne', css: 'top:-6px;right:-6px;cursor:ne-resize;' },
                { name: 'sw', css: 'bottom:-6px;left:-6px;cursor:sw-resize;' },
                { name: 'se', css: 'bottom:-6px;right:-6px;cursor:se-resize;' },
            ];

            corners.forEach(c => {
                const h = document.createElement('div');
                h.dataset.imgaction = 'resize-' + c.name;
                h.style.cssText = 'position:absolute;width:14px;height:14px;background:#2a5fbd;border:2px solid white;border-radius:3px;pointer-events:all;z-index:102;' + c.css;
                overlay.appendChild(h);
            });

            // Move handle — puri image ke upar
            const moveHandle = document.createElement('div');
            moveHandle.dataset.imgaction = 'move';
            moveHandle.style.cssText = 'position:absolute;top:0;left:0;width:100%;height:100%;cursor:move;pointer-events:all;z-index:101;';
            overlay.appendChild(moveHandle);

            editor.appendChild(overlay);
            updateOverlayPosition();
        };

        // Click handler
        const handleClick = (e) => {
            const target = e.target;
            // Agar handle par click hua toh kuch mat karo
            if (target.dataset && target.dataset.imgaction) return;
            if (target.dataset && target.dataset.overlay) return;

            if (target.tagName === 'IMG') {
                e.preventDefault();
                createOverlay(target);
            } else {
                clearImgSelection();
            }
        };

        // Mousedown on handles
        const handleMouseDown = (e) => {
            const act = e.target.dataset && e.target.dataset.imgaction;
            if (!act || !activeImg) return;
            e.preventDefault();
            e.stopPropagation();

            action = act;
            startX = e.clientX;
            startY = e.clientY;
            startW = activeImg.offsetWidth;
            startH = activeImg.offsetHeight;

            const cs = window.getComputedStyle(activeImg);
            startML = parseFloat(cs.marginLeft) || 0;
            startMT = parseFloat(cs.marginTop) || 0;

            document.addEventListener('mousemove', handleMouseMove);
            document.addEventListener('mouseup', handleMouseUp);
        };

        // Mouse move — resize ya move
        const handleMouseMove = (e) => {
            if (!action || !activeImg) return;
            e.preventDefault();
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;

            if (action.startsWith('resize-')) {
                const ratio = startW / startH;
                let newW;
                if (action === 'resize-se' || action === 'resize-ne') {
                    newW = startW + dx;
                } else {
                    newW = startW - dx;
                }
                if (newW < 40) newW = 40;
                const newH = newW / ratio;

                activeImg.style.width = newW + 'px';
                activeImg.style.height = newH + 'px';
                activeImg.style.maxWidth = 'none'; // max-width override taa ke resize kaam kare
                updateOverlayPosition();

            } else if (action === 'move') {
                activeImg.style.marginLeft = (startML + dx) + 'px';
                activeImg.style.marginTop = (startMT + dy) + 'px';
                updateOverlayPosition();
            }
        };

        const handleMouseUp = () => {
            action = null;
            document.removeEventListener('mousemove', handleMouseMove);
            document.removeEventListener('mouseup', handleMouseUp);
        };

        // Delete key se image delete
        const handleKeyDown = (e) => {
            if (activeImg && (e.key === 'Delete' || e.key === 'Backspace')) {
                e.preventDefault();
                activeImg.remove();
                clearImgSelection();
            }
        };

        editor.addEventListener('click', handleClick);
        editor.addEventListener('mousedown', handleMouseDown);
        document.addEventListener('keydown', handleKeyDown);

        return () => {
            editor.removeEventListener('click', handleClick);
            editor.removeEventListener('mousedown', handleMouseDown);
            document.removeEventListener('keydown', handleKeyDown);
            document.removeEventListener('mousemove', handleMouseMove);
            document.removeEventListener('mouseup', handleMouseUp);
        };
    }, []);
    // ========================================================

    // --- Page Size Logic ---
    const changePageSize = (event) => {
        const size = event.target.value;
        const editor = editorRef.current;
        if (!editor) return;
        if (size === "a4") {
            editor.style.width = "210mm";
            editor.style.margin = "20px auto";
        } else if (size === "legal") {
            editor.style.width = "216mm";
            editor.style.margin = "20px auto";
        } else {
            editor.style.width = "95%";
            editor.style.margin = "20px auto";
        }
    };

    const changeFontSize = (event) => {
        let size = event.target.value;
        if (!size) return;
        restoreSelection();
        document.execCommand('fontSize', false, 7);
        let fonts = document.querySelectorAll('font[size="7"]');
        fonts.forEach(font => {
            font.removeAttribute('size');
            font.style.fontSize = size;
        });
        saveSelection();
        event.target.value = "";
    };

    const changeFontFamily = (event) => {
        let font = event.target.value;
        if (!font) return;
        restoreSelection();
        document.execCommand('fontName', false, font);
        saveSelection();
    };

    // ✅ FIX: Image Upload — Direct DOM insertion (execCommand nahi, wo fail hota hai dialog ke baad)
    const handleImageUpload = (event) => {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            const editor = editorRef.current;
            if (!editor) return;

            // Image element banao
            const img = document.createElement('img');
            img.src = e.target.result;
            img.style.cssText = 'max-width:80%; height:auto; display:block; margin:10px auto; cursor:pointer;';

            // Editor ke end mein dalo
            const br = document.createElement('br');
            editor.appendChild(img);
            editor.appendChild(br);

            // Scroll down to show the new image
            img.scrollIntoView({ behavior: 'smooth', block: 'center' });
        };
        reader.readAsDataURL(file);
        // Reset input taa ke same file dubara upload ho sake
        event.target.value = '';
    };

    const insertShape = (event) => {
        const shapeType = event.target.value;
        let shapeHtml = "";
        if (shapeType === "square") shapeHtml = `<svg width="100" height="100"><rect width="100" height="100" style="fill:#3498db;" /></svg>`;
        else if (shapeType === "circle") shapeHtml = `<svg width="100" height="100"><circle cx="50" cy="50" r="50" style="fill:#e74c3c;" /></svg>`;
        else if (shapeType === "rectangle") shapeHtml = `<svg width="150" height="80"><rect width="150" height="80" style="fill:#2ecc71;" /></svg>`;

        if (shapeHtml) {
            restoreSelection();
            document.execCommand('insertHTML', false, `<br>${shapeHtml}<br>`);
            saveSelection();
        }
        event.target.value = "";
    };

    const addTable = () => {
        let rows = prompt("Rows?", "2");
        let cols = prompt("Cols?", "2");
        if (!rows || !cols) return;
        let table = `<br><table border="1" style="width:100%; border-collapse:collapse; text-align:center;">`;
        for (let i = 0; i < rows; i++) {
            table += "<tr>";
            for (let j = 0; j < cols; j++) table += "<td style='padding:8px;'>Data</td>";
            table += "</tr>";
        }
        table += "</table><br>";
        restoreSelection();
        document.execCommand('insertHTML', false, table);
        saveSelection();
    };

    const handleSave = () => {
        const editor = editorRef.current;
        // Save se pehle overlay hatao taa ke download clean ho
        const overlayEl = editor.querySelector('[data-overlay]');
        if (overlayEl) overlayEl.remove();
        // Image outlines bhi hatao
        editor.querySelectorAll('img').forEach(img => { img.style.outline = 'none'; });

        const content = editor.innerHTML;
        downloadFile(content, 'MyDocument.doc', 'application/msword');
    };

    return (
        <div className="editor-container" style={{ padding: '20px', display: 'flex', flexDirection: 'column', flexGrow: 1, backgroundColor: '#e9ecef' }}>
            <div className="editor-header">
                <h2>Word Professional Editor</h2>
                <div style={{ display: 'flex', gap: '10px' }}>
                    <button className="back-btn" onClick={onClose}>Back</button>
                    <button className="save-btn" onClick={handleSave} style={{ backgroundColor: '#2a5fbd', color: 'white' }}>Download (.doc)</button>
                </div>
            </div>

            {/* Toolbar */}
            <div className="toolbar" style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', padding: '12px', background: '#fff', borderBottom: '2px solid #ddd', zIndex: 10 }}>
                <button onMouseDown={keepFocus} onClick={() => execute('bold')}><b>B</b></button>
                <button onMouseDown={keepFocus} onClick={() => execute('italic')}><i>I</i></button>
                <button onMouseDown={keepFocus} onClick={() => execute('underline')}><u>U</u></button>
                <button onMouseDown={keepFocus} onClick={() => execute('strikeThrough')}><s>S</s></button>

                <select onChange={changeFontFamily} defaultValue="Arial">
                    <option value="Arial">Arial</option>
                    <option value="Times New Roman">Times New Roman</option>
                    <option value="Georgia">Georgia</option>
                    <option value="Courier New">Courier New</option>
                    <option value="Verdana">Verdana</option>
                    <option value="Comic Sans MS">Comic Sans</option>
                    <option value="Impact">Impact</option>
                    <option value="Garamond">Garamond</option>
                    <option value="Tahoma">Tahoma</option>
                    <option value="Trebuchet MS">Trebuchet</option>
                </select>

                <select onChange={changeFontSize}>
                    <option value="">Size</option>
                    <option value="12px">12</option>
                    <option value="14px">14</option>
                    <option value="16px">16</option>
                    <option value="20px">20</option>
                    <option value="24px">24</option>
                    <option value="36px">36</option>
                    <option value="48px">48</option>
                </select>

                <select onChange={changePageSize} style={{ border: '1px solid #2a5fbd', fontWeight: 'bold' }}>
                    <option value="a4">A4 Page</option>
                    <option value="legal">Legal</option>
                    <option value="full">Full Screen</option>
                </select>

                <div className="tool-group">
                    <label>Text:</label>
                    <input type="color" onMouseDown={keepFocus} onChange={(e) => execute('foreColor', e.target.value)} />
                </div>
                <div className="tool-group">
                    <label>Highlight:</label>
                    <input type="color" onMouseDown={keepFocus} onChange={(e) => execute('backColor', e.target.value)} defaultValue="#ffffff" />
                </div>

                <button onMouseDown={keepFocus} onClick={() => execute('justifyLeft')}><i className="fas fa-align-left"></i></button>
                <button onMouseDown={keepFocus} onClick={() => execute('justifyCenter')}><i className="fas fa-align-center"></i></button>
                <button onMouseDown={keepFocus} onClick={() => execute('justifyRight')}><i className="fas fa-align-right"></i></button>
                <button onMouseDown={keepFocus} onClick={() => execute('insertUnorderedList')}><i className="fas fa-list-ul"></i></button>
                <button onMouseDown={keepFocus} onClick={() => execute('insertOrderedList')}><i className="fas fa-list-ol"></i></button>

                <button onMouseDown={keepFocus} onClick={addTable}><i className="fas fa-table"></i> Table</button>
                <input type="file" accept="image/*" ref={fileInputRef} onChange={handleImageUpload} style={{ display: 'none' }} />
                <button onMouseDown={keepFocus} onClick={() => fileInputRef.current.click()}><i className="fas fa-image"></i> Image</button>

                <select onChange={insertShape}>
                    <option value="">Shapes</option>
                    <option value="square">Square</option>
                    <option value="circle">Circle</option>
                    <option value="rectangle">Rectangle</option>
                </select>

                <button onMouseDown={keepFocus} onClick={() => execute('undo')}><i className="fas fa-undo"></i></button>
                <button onMouseDown={keepFocus} onClick={() => execute('redo')}><i className="fas fa-redo"></i></button>
            </div>

            {/* Editor Area */}
            <div className="word-editor-scroll" style={{ overflowY: 'auto', flexGrow: 1, padding: '20px' }}>
                <div
                    id="editor-main"
                    ref={editorRef}
                    className="rich-editor"
                    contentEditable="true"
                    onMouseUp={saveSelection}
                    onKeyUp={saveSelection}
                    style={{
                        backgroundColor: 'white',
                        width: '210mm',
                        minHeight: '297mm',
                        margin: '0 auto',
                        padding: '20mm',
                        boxShadow: '0 0 15px rgba(0,0,0,0.2)',
                        outline: 'none',
                        fontSize: '16px',
                        lineHeight: '1.6',
                        position: 'relative',
                        backgroundImage: 'linear-gradient(to bottom, transparent 99%, #eee 100%)',
                        backgroundSize: '100% 297mm'
                    }}
                >
                    <div>Write Your Document Here....</div>
                </div>
            </div>
        </div>
    );
};

// --- Professional Presentation Editor ---
const PresentationEditor = ({ onClose }) => {
    const [slides, setSlides] = React.useState([
        { id: Date.now(), content: '<div style="text-align: center; margin-top: 15%; font-size: 36px; font-weight: bold;">Title Slide</div><div style="text-align: center; font-size: 20px; color: #555;">Click to edit subtitle</div>' }
    ]);
    const [activeSlide, setActiveSlide] = React.useState(0);

    const editorRef = React.useRef(null);
    const fileInputRef = React.useRef(null);

    // Jab active slide badle, toh editor canvas mein uska content dikhao
    React.useEffect(() => {
        if (editorRef.current) {
            editorRef.current.innerHTML = slides[activeSlide].content;
        }
    }, [activeSlide]);

    // Jab user canvas mein type kare, toh current slide ka state update ho
    const handleContentChange = () => {
        if (editorRef.current) {
            const newSlides = [...slides];
            // Overlay hatao content save karne se pehle
            const tempOverlay = editorRef.current.querySelector('[data-overlay]');
            if (tempOverlay) tempOverlay.remove();
            newSlides[activeSlide].content = editorRef.current.innerHTML;
            setSlides(newSlides);
        }
    };

    // ========================================================
    // ✅ IMAGE RESIZE & MOVE SYSTEM (PowerPoint ke liye)
    // ========================================================
    React.useEffect(() => {
        const editor = editorRef.current;
        if (!editor) return;

        let activeImg = null;
        let overlay = null;
        let action = null;
        let startX, startY, startW, startH, startML, startMT;

        const removeOverlay = () => {
            if (overlay && overlay.parentNode) {
                overlay.parentNode.removeChild(overlay);
            }
            overlay = null;
        };

        const clearImgSelection = () => {
            if (activeImg) {
                activeImg.style.outline = 'none';
                activeImg.style.cursor = '';
            }
            activeImg = null;
            removeOverlay();
        };

        const updateOverlayPosition = () => {
            if (!activeImg || !overlay) return;
            const editorRect = editor.getBoundingClientRect();
            const imgRect = activeImg.getBoundingClientRect();
            overlay.style.top = (imgRect.top - editorRect.top) + 'px';
            overlay.style.left = (imgRect.left - editorRect.left) + 'px';
            overlay.style.width = imgRect.width + 'px';
            overlay.style.height = imgRect.height + 'px';
        };

        const createOverlay = (img) => {
            clearImgSelection();
            activeImg = img;
            img.style.outline = '3px solid #8c19bf';

            overlay = document.createElement('div');
            overlay.setAttribute('data-overlay', 'true');
            overlay.style.cssText = 'position:absolute;z-index:100;box-sizing:border-box;pointer-events:none;';

            const corners = [
                { name: 'nw', css: 'top:-6px;left:-6px;cursor:nw-resize;' },
                { name: 'ne', css: 'top:-6px;right:-6px;cursor:ne-resize;' },
                { name: 'sw', css: 'bottom:-6px;left:-6px;cursor:sw-resize;' },
                { name: 'se', css: 'bottom:-6px;right:-6px;cursor:se-resize;' },
            ];

            corners.forEach(c => {
                const h = document.createElement('div');
                h.dataset.imgaction = 'resize-' + c.name;
                h.style.cssText = 'position:absolute;width:14px;height:14px;background:#8c19bf;border:2px solid white;border-radius:3px;pointer-events:all;z-index:102;' + c.css;
                overlay.appendChild(h);
            });

            const moveHandle = document.createElement('div');
            moveHandle.dataset.imgaction = 'move';
            moveHandle.style.cssText = 'position:absolute;top:0;left:0;width:100%;height:100%;cursor:move;pointer-events:all;z-index:101;';
            overlay.appendChild(moveHandle);

            editor.appendChild(overlay);
            updateOverlayPosition();
        };

        const handleClick = (e) => {
            const target = e.target;
            if (target.dataset && target.dataset.imgaction) return;
            if (target.dataset && target.dataset.overlay) return;

            if (target.tagName === 'IMG') {
                e.preventDefault();
                createOverlay(target);
            } else {
                clearImgSelection();
            }
        };

        const handleMouseDown = (e) => {
            const act = e.target.dataset && e.target.dataset.imgaction;
            if (!act || !activeImg) return;
            e.preventDefault();
            e.stopPropagation();

            action = act;
            startX = e.clientX;
            startY = e.clientY;
            startW = activeImg.offsetWidth;
            startH = activeImg.offsetHeight;

            const cs = window.getComputedStyle(activeImg);
            startML = parseFloat(cs.marginLeft) || 0;
            startMT = parseFloat(cs.marginTop) || 0;

            document.addEventListener('mousemove', handleMouseMove);
            document.addEventListener('mouseup', handleMouseUp);
        };

        const handleMouseMove = (e) => {
            if (!action || !activeImg) return;
            e.preventDefault();
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;

            if (action.startsWith('resize-')) {
                const ratio = startW / startH;
                let newW;
                if (action === 'resize-se' || action === 'resize-ne') {
                    newW = startW + dx;
                } else {
                    newW = startW - dx;
                }
                if (newW < 40) newW = 40;
                const newH = newW / ratio;

                activeImg.style.width = newW + 'px';
                activeImg.style.height = newH + 'px';
                activeImg.style.maxWidth = 'none';
                activeImg.style.maxHeight = 'none';
                updateOverlayPosition();

            } else if (action === 'move') {
                activeImg.style.marginLeft = (startML + dx) + 'px';
                activeImg.style.marginTop = (startMT + dy) + 'px';
                updateOverlayPosition();
            }
        };

        const handleMouseUp = () => {
            action = null;
            document.removeEventListener('mousemove', handleMouseMove);
            document.removeEventListener('mouseup', handleMouseUp);
        };

        const handleKeyDown = (e) => {
            if (activeImg && (e.key === 'Delete' || e.key === 'Backspace')) {
                e.preventDefault();
                activeImg.remove();
                clearImgSelection();
            }
        };

        editor.addEventListener('click', handleClick);
        editor.addEventListener('mousedown', handleMouseDown);
        document.addEventListener('keydown', handleKeyDown);

        return () => {
            editor.removeEventListener('click', handleClick);
            editor.removeEventListener('mousedown', handleMouseDown);
            document.removeEventListener('keydown', handleKeyDown);
            document.removeEventListener('mousemove', handleMouseMove);
            document.removeEventListener('mouseup', handleMouseUp);
        };
    }, [activeSlide]);
    // ========================================================

    // --- Slide Management ---
    const addSlide = () => {
        const newSlide = {
            id: Date.now(),
            content: '<div style="font-size: 28px; font-weight: bold; margin-bottom: 20px;">New Slide</div><div>Click to add text</div>'
        };
        setSlides([...slides, newSlide]);
        setActiveSlide(slides.length);
    };

    const deleteSlide = (index, e) => {
        e.stopPropagation();
        if (slides.length === 1) {
            alert("Aap aakhri slide delete nahi kar sakte!");
            return;
        }
        const newSlides = slides.filter((_, i) => i !== index);
        setSlides(newSlides);
        setActiveSlide(Math.max(0, index - 1));
    };

    // --- Toolbar Formatting ---
    const execute = (cmd, val = null) => {
        document.execCommand(cmd, false, val);
        handleContentChange();
    };

    const keepFocus = (e) => e.preventDefault();

    const changeFontSize = (event) => {
        let size = event.target.value;
        if (!size) return;
        document.execCommand('fontSize', false, 7);
        let fonts = editorRef.current.querySelectorAll('font[size="7"]');
        fonts.forEach(font => {
            font.removeAttribute('size');
            font.style.fontSize = size;
        });
        handleContentChange();
        event.target.value = "";
    };

    // ✅ FIX: Direct DOM insertion for image upload (execCommand fail hota tha)
    const handleImageUpload = (event) => {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            const editor = editorRef.current;
            if (!editor) return;

            const img = document.createElement('img');
            img.src = e.target.result;
            img.style.cssText = 'max-width:80%; max-height:300px; display:block; margin:10px auto; cursor:pointer;';
            editor.appendChild(img);
            handleContentChange();
        };
        reader.readAsDataURL(file);
        event.target.value = '';
    };

    const insertShape = (event) => {
        const shapeType = event.target.value;
        let shapeHtml = "";
        if (shapeType === "square") shapeHtml = `<svg width="100" height="100"><rect width="100" height="100" style="fill:#3498db;" /></svg>`;
        else if (shapeType === "circle") shapeHtml = `<svg width="100" height="100"><circle cx="50" cy="50" r="50" style="fill:#e74c3c;" /></svg>`;
        else if (shapeType === "rectangle") shapeHtml = `<svg width="150" height="80"><rect width="150" height="80" style="fill:#2ecc71;" /></svg>`;

        if (shapeHtml) {
            execute('insertHTML', shapeHtml);
        }
        event.target.value = "";
    };

    // --- Save/Download ---
    const handleSave = () => {
        // Save se pehle overlay hatao
        if (editorRef.current) {
            const ov = editorRef.current.querySelector('[data-overlay]');
            if (ov) ov.remove();
            editorRef.current.querySelectorAll('img').forEach(img => { img.style.outline = 'none'; });
            handleContentChange();
        }

        let htmlContent = `
            <html><head><title>My Presentation</title>
            <style>
                body { background: #333; display: flex; flex-direction: column; align-items: center; padding: 40px; font-family: sans-serif; }
                .slide { width: 960px; height: 540px; background: white; margin-bottom: 30px; padding: 40px; box-shadow: 0 4px 15px rgba(0,0,0,0.5); overflow: hidden; position: relative; }
                @media print { body { background: white; padding: 0; } .slide { box-shadow: none; margin: 0; page-break-after: always; } }
            </style>
            </head><body>
        `;
        slides.forEach((s) => {
            htmlContent += `<div class="slide">${s.content}</div>`;
        });
        htmlContent += `</body></html>`;
        downloadFile(htmlContent, 'MyPresentation.html', 'text/html');
    };

    return (
        <div className="editor-container" style={{ display: 'flex', flexDirection: 'column', flexGrow: 1, backgroundColor: '#f3f4f6', height: '100vh' }}>

            {/* Header */}
            <div className="editor-header" style={{ padding: '15px 20px', background: 'white', borderBottom: '1px solid #ddd', display: 'flex', justifyContent: 'space-between' }}>
                <h2>PowerPoint Editor</h2>
                <div style={{ display: 'flex', gap: '10px' }}>
                    <button className="back-btn" onClick={onClose} style={{ padding: '8px 15px', cursor: 'pointer' }}>Back</button>
                    <button className="save-btn pres-btn" onClick={handleSave} style={{ backgroundColor: '#8c19bf', color: 'white', padding: '8px 15px', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Download (.html)</button>
                </div>
            </div>

            {/* Toolbar */}
            <div className="toolbar" style={{ display: 'flex', flexWrap: 'wrap', gap: '10px', padding: '10px 20px', background: '#fff', borderBottom: '1px solid #ddd', zIndex: 10 }}>
                <button onMouseDown={keepFocus} onClick={() => execute('bold')} style={{ padding: '5px 10px', cursor: 'pointer' }}><b>B</b></button>
                <button onMouseDown={keepFocus} onClick={() => execute('italic')} style={{ padding: '5px 10px', cursor: 'pointer' }}><i>I</i></button>
                <button onMouseDown={keepFocus} onClick={() => execute('underline')} style={{ padding: '5px 10px', cursor: 'pointer' }}><u>U</u></button>

                <select onChange={changeFontSize} style={{ padding: '5px' }}>
                    <option value="">Size</option>
                    <option value="16px">16</option>
                    <option value="24px">24</option>
                    <option value="36px">36 (Title)</option>
                    <option value="48px">48</option>
                    <option value="72px">72</option>
                </select>

                <div style={{ display: 'flex', alignItems: 'center', gap: '5px' }}>
                    <label style={{ fontSize: '12px' }}>Color:</label>
                    <input type="color" onMouseDown={keepFocus} onChange={(e) => execute('foreColor', e.target.value)} />
                </div>

                <button onMouseDown={keepFocus} onClick={() => execute('justifyLeft')}><i className="fas fa-align-left"></i></button>
                <button onMouseDown={keepFocus} onClick={() => execute('justifyCenter')}><i className="fas fa-align-center"></i></button>
                <button onMouseDown={keepFocus} onClick={() => execute('insertUnorderedList')}><i className="fas fa-list-ul"></i></button>

                <input type="file" accept="image/*" ref={fileInputRef} onChange={handleImageUpload} style={{ display: 'none' }} />
                <button onMouseDown={keepFocus} onClick={() => fileInputRef.current.click()}><i className="fas fa-image"></i> Image</button>

                <select onChange={insertShape} style={{ padding: '5px' }}>
                    <option value="">Add Shape</option>
                    <option value="square">Square</option>
                    <option value="circle">Circle</option>
                    <option value="rectangle">Rectangle</option>
                </select>
            </div>

            {/* Workspace Area: Sidebar + Canvas */}
            <div style={{ display: 'flex', flexGrow: 1, overflow: 'hidden' }}>

                {/* Left Sidebar: Slide Thumbnails */}
                <div style={{ width: '250px', background: '#fff', borderRight: '1px solid #ddd', overflowY: 'auto', padding: '10px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
                    <button onClick={addSlide} style={{ padding: '10px', background: '#8c19bf', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
                        + Add New Slide
                    </button>
                    <hr style={{ border: 'none', borderTop: '1px solid #eee' }} />

                    {slides.map((slide, index) => (
                        <div
                            key={slide.id}
                            onClick={() => setActiveSlide(index)}
                            style={{
                                height: '120px',
                                border: activeSlide === index ? '2px solid #8c19bf' : '1px solid #ccc',
                                borderRadius: '4px',
                                background: 'white',
                                padding: '5px',
                                cursor: 'pointer',
                                position: 'relative',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                fontSize: '10px',
                                overflow: 'hidden'
                            }}
                        >
                            <span style={{ position: 'absolute', top: '5px', left: '5px', background: '#ddd', padding: '2px 6px', borderRadius: '3px', fontWeight: 'bold' }}>{index + 1}</span>
                            <button
                                onClick={(e) => deleteSlide(index, e)}
                                style={{ position: 'absolute', top: '5px', right: '5px', background: '#e74c3c', color: 'white', border: 'none', borderRadius: '3px', cursor: 'pointer', padding: '2px 5px' }}
                                title="Delete Slide"
                            >X</button>
                            <div dangerouslySetInnerHTML={{ __html: slide.content }} style={{ transform: 'scale(0.2)', transformOrigin: 'center center', width: '960px', height: '540px', pointerEvents: 'none' }} />
                        </div>
                    ))}
                </div>

                {/* Right Area: Active Slide Canvas */}
                <div style={{ flexGrow: 1, background: '#e9ecef', padding: '40px', display: 'flex', justifyContent: 'center', alignItems: 'center', overflow: 'auto' }}>
                    <div
                        ref={editorRef}
                        contentEditable="true"
                        onKeyUp={handleContentChange}
                        onMouseUp={handleContentChange}
                        style={{
                            width: '960px',
                            height: '540px',
                            background: 'white',
                            boxShadow: '0 4px 15px rgba(0,0,0,0.1)',
                            padding: '40px',
                            outline: 'none',
                            position: 'relative',
                            overflow: 'hidden'
                        }}
                    >
                    </div>
                </div>

            </div>
        </div>
    );
};

// Simplified Sheet Editor (basic grid with CSV download)
// --- Professional Excel-Style Sheet Editor ---
const SheetEditor = ({ onClose }) => {
    const ROWS = 100;
    const COLS = 26;

    const getColName = (n) => String.fromCharCode(65 + n);

    const createInitialGrid = () => {
        let initialGrid = [];
        for (let r = 0; r < ROWS; r++) {
            let row = [];
            for (let c = 0; c < COLS; c++) {
                row.push({
                    id: `${getColName(c)}${r + 1}`,
                    raw: '',
                    computed: '',
                    bold: false,
                    italic: false,
                    underline: false,
                    align: 'left',
                    color: '#000000',
                    bg: '#ffffff'
                });
            }
            initialGrid.push(row);
        }
        return initialGrid;
    };

    const [sheets, setSheets] = React.useState([
        { id: 1, name: 'Sheet 1', grid: createInitialGrid() },
        { id: 2, name: 'Sheet 2', grid: createInitialGrid() },
        { id: 3, name: 'Sheet 3', grid: createInitialGrid() }
    ]);
    const [activeSheetIndex, setActiveSheetIndex] = React.useState(0);

    const [activeCell, setActiveCell] = React.useState({ r: 0, c: 0 });
    const [formulaBar, setFormulaBar] = React.useState('');

    // ✅ NEW: Selection range state for Shift+Click multi-select
    const [selectionStart, setSelectionStart] = React.useState(null); // {r, c}
    const [selectionEnd, setSelectionEnd] = React.useState(null);     // {r, c}

    const currentGrid = sheets[activeSheetIndex].grid;

    const updateCurrentGrid = (newGrid) => {
        const newSheets = [...sheets];
        newSheets[activeSheetIndex].grid = newGrid;
        setSheets(newSheets);
    };

    // ✅ Helper: Check if a cell is inside the selection range
    const isCellSelected = (r, c) => {
        if (!selectionStart || !selectionEnd) {
            return activeCell.r === r && activeCell.c === c;
        }
        const minR = Math.min(selectionStart.r, selectionEnd.r);
        const maxR = Math.max(selectionStart.r, selectionEnd.r);
        const minC = Math.min(selectionStart.c, selectionEnd.c);
        const maxC = Math.max(selectionStart.c, selectionEnd.c);
        return r >= minR && r <= maxR && c >= minC && c <= maxC;
    };

    // ✅ Helper: Check if multi-selection is active (more than 1 cell)
    const hasMultiSelection = () => {
        if (!selectionStart || !selectionEnd) return false;
        return selectionStart.r !== selectionEnd.r || selectionStart.c !== selectionEnd.c;
    };

    // --- Formula Evaluator ---
    const evaluateFormula = (formula, currentGridState) => {
        if (!formula.startsWith('=')) return formula;

        try {
            let expr = formula.substring(1).toUpperCase();

            if (expr.startsWith('SUM(') && expr.endsWith(')')) {
                const range = expr.replace('SUM(', '').replace(')', '').split(':');
                if (range.length === 2) {
                    let sum = 0;
                    currentGridState.forEach(row => {
                        row.forEach(cell => {
                            if (cell.id >= range[0] && cell.id <= range[1]) {
                                sum += parseFloat(cell.computed) || 0;
                            }
                        });
                    });
                    return sum;
                }
            }

            expr = expr.replace(/[A-Z][0-9]+/g, (match) => {
                let val = 0;
                currentGridState.forEach(row => {
                    row.forEach(cell => {
                        if (cell.id === match) val = parseFloat(cell.computed) || 0;
                    });
                });
                return val;
            });

            return Function(`'use strict'; return (${expr})`)();
        } catch (e) {
            return '#ERROR!';
        }
    };

    const recalculateGrid = (newGrid) => {
        let updatedGrid = [...newGrid];
        for (let r = 0; r < ROWS; r++) {
            for (let c = 0; c < COLS; c++) {
                if (updatedGrid[r][c].raw.startsWith('=')) {
                    updatedGrid[r][c].computed = evaluateFormula(updatedGrid[r][c].raw, newGrid);
                } else {
                    updatedGrid[r][c].computed = updatedGrid[r][c].raw;
                }
            }
        }
        return updatedGrid;
    };

    const handleCellChange = (r, c, value) => {
        let newGrid = [...currentGrid];
        newGrid[r][c].raw = value;
        newGrid = recalculateGrid(newGrid);
        updateCurrentGrid(newGrid);
        setFormulaBar(value);
    };

    const handleFormulaBarChange = (e) => {
        const val = e.target.value;
        setFormulaBar(val);
        handleCellChange(activeCell.r, activeCell.c, val);
    };

    // ✅ NEW: Cell Click handler — Shift+Click se range select hoga
    const handleCellClick = (r, c, e) => {
        if (e.shiftKey && selectionStart) {
            // Shift+Click: Range extend karo
            setSelectionEnd({ r, c });
            setActiveCell({ r, c });
            setFormulaBar(currentGrid[r][c].raw);
        } else {
            // Normal click: Single cell select
            setActiveCell({ r, c });
            setSelectionStart({ r, c });
            setSelectionEnd({ r, c });
            setFormulaBar(currentGrid[r][c].raw);
        }
    };

    // ✅ NEW: Keyboard Navigation — Enter, Tab, Arrow keys
    const handleCellKeyDown = (e, r, c) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            const newR = Math.min(r + 1, ROWS - 1);
            setActiveCell({ r: newR, c });
            setSelectionStart({ r: newR, c });
            setSelectionEnd({ r: newR, c });
            setFormulaBar(currentGrid[newR][c].raw);
            const nextInput = document.querySelector(`[data-cell="${newR}-${c}"]`);
            if (nextInput) nextInput.focus();

        } else if (e.key === 'Tab') {
            e.preventDefault();
            const newC = e.shiftKey ? Math.max(c - 1, 0) : Math.min(c + 1, COLS - 1);
            setActiveCell({ r, c: newC });
            setSelectionStart({ r, c: newC });
            setSelectionEnd({ r, c: newC });
            setFormulaBar(currentGrid[r][newC].raw);
            const nextInput = document.querySelector(`[data-cell="${r}-${newC}"]`);
            if (nextInput) nextInput.focus();

        } else if (e.key === 'ArrowDown') {
            e.preventDefault();
            const newR = Math.min(r + 1, ROWS - 1);
            if (e.shiftKey) {
                setSelectionEnd({ r: newR, c });
            } else {
                setSelectionStart({ r: newR, c });
                setSelectionEnd({ r: newR, c });
            }
            setActiveCell({ r: newR, c });
            setFormulaBar(currentGrid[newR][c].raw);
            const nextInput = document.querySelector(`[data-cell="${newR}-${c}"]`);
            if (nextInput) nextInput.focus();

        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            const newR = Math.max(r - 1, 0);
            if (e.shiftKey) {
                setSelectionEnd({ r: newR, c });
            } else {
                setSelectionStart({ r: newR, c });
                setSelectionEnd({ r: newR, c });
            }
            setActiveCell({ r: newR, c });
            setFormulaBar(currentGrid[newR][c].raw);
            const nextInput = document.querySelector(`[data-cell="${newR}-${c}"]`);
            if (nextInput) nextInput.focus();

        } else if (e.key === 'ArrowRight') {
            e.preventDefault();
            const newC = Math.min(c + 1, COLS - 1);
            if (e.shiftKey) {
                setSelectionEnd({ r, c: newC });
            } else {
                setSelectionStart({ r, c: newC });
                setSelectionEnd({ r, c: newC });
            }
            setActiveCell({ r, c: newC });
            setFormulaBar(currentGrid[r][newC].raw);
            const nextInput = document.querySelector(`[data-cell="${r}-${newC}"]`);
            if (nextInput) nextInput.focus();

        } else if (e.key === 'ArrowLeft') {
            e.preventDefault();
            const newC = Math.max(c - 1, 0);
            if (e.shiftKey) {
                setSelectionEnd({ r, c: newC });
            } else {
                setSelectionStart({ r, c: newC });
                setSelectionEnd({ r, c: newC });
            }
            setActiveCell({ r, c: newC });
            setFormulaBar(currentGrid[r][newC].raw);
            const nextInput = document.querySelector(`[data-cell="${r}-${newC}"]`);
            if (nextInput) nextInput.focus();
        }
    };

    // ✅ NEW: Formatting apply to ALL selected cells
    const applyFormat = (formatType, value = null) => {
        let newGrid = [...currentGrid];

        const minR = selectionStart && selectionEnd ? Math.min(selectionStart.r, selectionEnd.r) : activeCell.r;
        const maxR = selectionStart && selectionEnd ? Math.max(selectionStart.r, selectionEnd.r) : activeCell.r;
        const minC = selectionStart && selectionEnd ? Math.min(selectionStart.c, selectionEnd.c) : activeCell.c;
        const maxC = selectionStart && selectionEnd ? Math.max(selectionStart.c, selectionEnd.c) : activeCell.c;

        for (let r = minR; r <= maxR; r++) {
            for (let c = minC; c <= maxC; c++) {
                let cell = newGrid[r][c];
                if (formatType === 'bold') cell.bold = !cell.bold;
                if (formatType === 'italic') cell.italic = !cell.italic;
                if (formatType === 'underline') cell.underline = !cell.underline;
                if (formatType === 'align') cell.align = value;
                if (formatType === 'color') cell.color = value;
                if (formatType === 'bg') cell.bg = value;
            }
        }
        updateCurrentGrid(newGrid);
    };

    const handleSave = () => {
        const csvString = currentGrid.map(row =>
            row.map(cell => `"${cell.computed}"`).join(',')
        ).join('\n');
        downloadFile(csvString, `${sheets[activeSheetIndex].name}.csv`, 'text/csv');
    };

    const activeCellData = currentGrid[activeCell.r][activeCell.c];

    // ✅ Selection info text
    const getSelectionInfo = () => {
        if (!hasMultiSelection()) return activeCellData.id;
        const minR = Math.min(selectionStart.r, selectionEnd.r);
        const maxR = Math.max(selectionStart.r, selectionEnd.r);
        const minC = Math.min(selectionStart.c, selectionEnd.c);
        const maxC = Math.max(selectionStart.c, selectionEnd.c);
        const count = (maxR - minR + 1) * (maxC - minC + 1);
        return `${getColName(minC)}${minR + 1}:${getColName(maxC)}${maxR + 1} (${count} cells)`;
    };

    return (
        <div className="editor-container" style={{ display: 'flex', flexDirection: 'column', flexGrow: 1, backgroundColor: '#f3f4f6', height: '100vh', overflow: 'hidden' }}>

            {/* Header */}
            <div className="editor-header" style={{ padding: '15px 20px', background: 'white', borderBottom: '1px solid #ddd', display: 'flex', justifyContent: 'space-between' }}>
                <h2>Excel-Style Editor</h2>
                <div style={{ display: 'flex', gap: '10px' }}>
                    <button className="back-btn" onClick={onClose} style={{ padding: '8px 15px', cursor: 'pointer' }}>Back</button>
                    <button className="save-btn sheet-btn" onClick={handleSave} style={{ backgroundColor: '#0c9e58', color: 'white', padding: '8px 15px', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Download Current Sheet</button>
                </div>
            </div>

            {/* Formatting Toolbar */}
            <div className="toolbar" style={{ display: 'flex', flexWrap: 'wrap', gap: '10px', padding: '10px 20px', background: '#fff', borderBottom: '1px solid #ddd' }}>
                <button onClick={() => applyFormat('bold')} style={{ fontWeight: 'bold', background: activeCellData.bold ? '#ddd' : '#fff' }}>B</button>
                <button onClick={() => applyFormat('italic')} style={{ fontStyle: 'italic', background: activeCellData.italic ? '#ddd' : '#fff' }}>I</button>
                <button onClick={() => applyFormat('underline')} style={{ textDecoration: 'underline', background: activeCellData.underline ? '#ddd' : '#fff' }}>U</button>

                <div style={{ borderLeft: '1px solid #ccc', margin: '0 5px' }}></div>

                <button onClick={() => applyFormat('align', 'left')}><i className="fas fa-align-left"></i></button>
                <button onClick={() => applyFormat('align', 'center')}><i className="fas fa-align-center"></i></button>
                <button onClick={() => applyFormat('align', 'right')}><i className="fas fa-align-right"></i></button>

                <div style={{ borderLeft: '1px solid #ccc', margin: '0 5px' }}></div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '5px' }}>
                    <label style={{ fontSize: '12px' }}>Text:</label>
                    <input type="color" value={activeCellData.color} onChange={(e) => applyFormat('color', e.target.value)} />
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '5px' }}>
                    <label style={{ fontSize: '12px' }}>Fill:</label>
                    <input type="color" value={activeCellData.bg} onChange={(e) => applyFormat('bg', e.target.value)} />
                </div>
            </div>

            {/* Formula Bar */}
            <div style={{ display: 'flex', padding: '10px 20px', background: '#fff', borderBottom: '1px solid #ddd', alignItems: 'center', gap: '10px' }}>
                <div style={{ minWidth: '80px', fontWeight: 'bold', textAlign: 'center', background: '#f3f4f6', padding: '5px', border: '1px solid #ccc', fontSize: '12px' }}>
                    {getSelectionInfo()}
                </div>
                <div style={{ fontWeight: 'bold', color: '#0c9e58', fontStyle: 'italic' }}>fx</div>
                <input
                    type="text"
                    value={formulaBar}
                    onChange={handleFormulaBarChange}
                    style={{ flexGrow: 1, padding: '8px', border: '1px solid #ccc', outline: 'none', fontFamily: 'monospace' }}
                    placeholder="Enter value or formula (e.g., =A1+B2)"
                />
            </div>

            {/* Spreadsheet Grid */}
            <div style={{ overflow: 'auto', flexGrow: 1, background: '#e9ecef', padding: '0' }}>
                <table style={{ borderCollapse: 'collapse', background: 'white', width: '100%' }}>
                    <thead style={{ position: 'sticky', top: 0, zIndex: 5 }}>
                        <tr>
                            <th style={{ width: '40px', background: '#f3f4f6', border: '1px solid #ccc', position: 'sticky', left: 0, zIndex: 6 }}></th>
                            {Array.from({ length: COLS }).map((_, i) => (
                                <th key={i} style={{ minWidth: '100px', background: '#f3f4f6', border: '1px solid #ccc', padding: '5px', textAlign: 'center', fontWeight: 'normal', color: '#555' }}>
                                    {getColName(i)}
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {currentGrid.map((row, r) => (
                            <tr key={r}>
                                <td style={{ background: '#f3f4f6', border: '1px solid #ccc', textAlign: 'center', fontWeight: 'normal', color: '#555', position: 'sticky', left: 0, zIndex: 4 }}>
                                    {r + 1}
                                </td>
                                {row.map((cell, c) => {
                                    const isActive = activeCell.r === r && activeCell.c === c;
                                    const isSelected = isCellSelected(r, c);
                                    return (
                                        <td
                                            key={c}
                                            onClick={(e) => handleCellClick(r, c, e)}
                                            style={{
                                                border: isActive ? '2px solid #0c9e58' : '1px solid #ddd',
                                                padding: 0,
                                                margin: 0,
                                                background: isSelected && !isActive ? '#d4edda' : cell.bg,
                                                minWidth: '100px'
                                            }}
                                        >
                                            <input
                                                type="text"
                                                data-cell={`${r}-${c}`}
                                                value={isActive ? cell.raw : cell.computed}
                                                onChange={(e) => handleCellChange(r, c, e.target.value)}
                                                onFocus={() => {
                                                    setActiveCell({ r, c });
                                                    setFormulaBar(cell.raw);
                                                    // Selection logic sirf handleCellClick mein hai
                                                    // onFocus mein selection reset NAHI karenge
                                                }}
                                                onKeyDown={(e) => handleCellKeyDown(e, r, c)}
                                                style={{
                                                    width: '100%',
                                                    height: '100%',
                                                    border: 'none',
                                                    outline: 'none',
                                                    padding: '8px',
                                                    background: 'transparent',
                                                    fontWeight: cell.bold ? 'bold' : 'normal',
                                                    fontStyle: cell.italic ? 'italic' : 'normal',
                                                    textDecoration: cell.underline ? 'underline' : 'none',
                                                    textAlign: cell.align,
                                                    color: cell.color,
                                                    cursor: 'cell'
                                                }}
                                            />
                                        </td>
                                    );
                                })}
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

            {/* Bottom Tabs for Multiple Sheets */}
            <div style={{ display: 'flex', background: '#f3f4f6', borderTop: '1px solid #ccc', padding: '5px 10px', alignItems: 'flex-end', gap: '5px' }}>
                {sheets.map((sheet, index) => (
                    <div
                        key={sheet.id}
                        onClick={() => {
                            setActiveSheetIndex(index);
                            setActiveCell({ r: 0, c: 0 });
                            setSelectionStart(null);
                            setSelectionEnd(null);
                            setFormulaBar('');
                        }}
                        style={{
                            padding: '8px 20px',
                            background: activeSheetIndex === index ? 'white' : '#e5e7eb',
                            border: '1px solid #ccc',
                            borderBottom: activeSheetIndex === index ? 'none' : '1px solid #ccc',
                            borderRadius: '5px 5px 0 0',
                            cursor: 'pointer',
                            fontWeight: activeSheetIndex === index ? 'bold' : 'normal',
                            color: activeSheetIndex === index ? '#0c9e58' : '#333'
                        }}
                    >
                        {sheet.name}
                    </div>
                ))}

                <button
                    onClick={() => {
                        const newSheet = { id: Date.now(), name: `Sheet ${sheets.length + 1}`, grid: createInitialGrid() };
                        setSheets([...sheets, newSheet]);
                        setActiveSheetIndex(sheets.length);
                        setActiveCell({ r: 0, c: 0 });
                        setSelectionStart(null);
                        setSelectionEnd(null);
                        setFormulaBar('');
                    }}
                    style={{
                        padding: '5px 12px',
                        background: 'white',
                        border: '1px solid #ccc',
                        borderRadius: '15px',
                        cursor: 'pointer',
                        fontWeight: 'bold',
                        color: '#555',
                        marginLeft: '10px',
                        marginBottom: '3px'
                    }}
                    title="Add New Sheet"
                >
                    +
                </button>
            </div>
        </div>
    );
};

// --- Legal Pages Components ---
const PrivacyPolicy = () => {
    return (
        <div style={{ padding: '40px 10%', flexGrow: 1, backgroundColor: '#fff', color: '#333', lineHeight: '1.6' }}>
            <h1 style={{ marginBottom: '20px', color: '#1a2033' }}>Privacy Policy</h1>
            <p style={{ marginBottom: '15px' }}><strong>Effective Date:</strong> April 2026</p>
            <h3 style={{ marginTop: '20px', color: '#2a5fbd' }}>1. Information We Collect</h3>
            <p style={{ marginBottom: '15px' }}>Office Lite operates primarily as a client-side web application. We do not require you to create an account, nor do we collect, store, or process any personal data, documents, or files on our servers. All documents, presentations, and spreadsheets you create are processed locally within your browser.</p>
            <h3 style={{ marginTop: '20px', color: '#2a5fbd' }}>2. How We Use Your Data</h3>
            <p style={{ marginBottom: '15px' }}>Since your data remains on your local device, we do not use your information for analytics, marketing, or any other tracking purposes. The files you download are generated directly by your browser.</p>
            <h3 style={{ marginTop: '20px', color: '#2a5fbd' }}>3. Third-Party Services</h3>
            <p style={{ marginBottom: '15px' }}>We do not share any data with third-party services. However, if you use browser extensions or third-party tools alongside Office Lite, their respective privacy policies will apply.</p>
            <h3 style={{ marginTop: '20px', color: '#2a5fbd' }}>4. Changes to This Policy</h3>
            <p style={{ marginBottom: '15px' }}>We may update this Privacy Policy from time to time. Any changes will be reflected on this page with an updated effective date.</p>
            <p style={{ marginTop: '30px', fontStyle: 'italic' }}>If you have any questions, please contact us at resumeprohub1@gmail.com.</p>
        </div>
    );
};

const TermsAndConditions = () => {
    return (
        <div style={{ padding: '40px 10%', flexGrow: 1, backgroundColor: '#fff', color: '#333', lineHeight: '1.6' }}>
            <h1 style={{ marginBottom: '20px', color: '#1a2033' }}>Terms and Conditions</h1>
            <p style={{ marginBottom: '15px' }}><strong>Effective Date:</strong> April 2026</p>
            <h3 style={{ marginTop: '20px', color: '#2a5fbd' }}>1. Acceptance of Terms</h3>
            <p style={{ marginBottom: '15px' }}>By accessing and using Office Lite, you accept and agree to be bound by the terms and provision of this agreement. If you do not agree to abide by these terms, please do not use our service.</p>
            <h3 style={{ marginTop: '20px', color: '#2a5fbd' }}>2. Use of Service</h3>
            <p style={{ marginBottom: '15px' }}>Office Lite provides tools for creating and editing documents, presentations, and spreadsheets. You agree to use these tools only for lawful purposes. You are solely responsible for the content you create and download using our platform.</p>
            <h3 style={{ marginTop: '20px', color: '#2a5fbd' }}>3. Disclaimer of Warranties</h3>
            <p style={{ marginBottom: '15px' }}>The service is provided on an "as is" and "as available" basis. Office Lite makes no warranties, expressed or implied, and hereby disclaims all warranties, including without limitation, implied warranties of merchantability and fitness for a particular purpose.</p>
            <h3 style={{ marginTop: '20px', color: '#2a5fbd' }}>4. Limitation of Liability</h3>
            <p style={{ marginBottom: '15px' }}>In no event shall Office Lite be liable for any data loss, damages, or issues arising out of the use or inability to use the materials on our website.</p>
        </div>
    );
};

// --- Contact Us Component ---
const ContactUs = () => {
    return (
        <div style={{ padding: '40px 10%', flexGrow: 1, backgroundColor: '#e9ecef', color: '#333', display: 'flex', justifyContent: 'center', alignItems: 'center' }}>
            <div style={{ width: '100%', maxWidth: '600px', backgroundColor: '#fff', padding: '40px', borderRadius: '12px', boxShadow: '0 8px 24px rgba(0,0,0,0.1)' }}>
                <h2 style={{ textAlign: 'center', color: '#1a2033', marginBottom: '10px', fontSize: '2rem' }}>Contact Us</h2>
                <p style={{ textAlign: 'center', color: '#6c757d', marginBottom: '30px' }}>Have a question, feedback, or need support? Drop us a message!</p>

                {/* FormSubmit.co ka use kar ke direct email bhejna */}
                <form action="https://formsubmit.co/resumeprohub1@gmail.com" method="POST" style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
                    {/* Security aur UX ke liye hidden fields */}
                    <input type="hidden" name="_subject" value="New Contact Message from Office Lite!" />
                    <input type="hidden" name="_captcha" value="false" />

                    <div>
                        <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold', color: '#1a2033' }}>Your Name</label>
                        <input type="text" name="name" required placeholder="Name" style={{ width: '100%', padding: '12px', borderRadius: '6px', border: '1px solid #ccc', outline: 'none', fontSize: '1rem' }} />
                    </div>

                    <div>
                        <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold', color: '#1a2033' }}>Your Email</label>
                        <input type="email" name="email" required placeholder="Email" style={{ width: '100%', padding: '12px', borderRadius: '6px', border: '1px solid #ccc', outline: 'none', fontSize: '1rem' }} />
                    </div>

                    <div>
                        <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold', color: '#1a2033' }}>Message</label>
                        <textarea name="message" required placeholder="How can we help you?" rows="5" style={{ width: '100%', padding: '12px', borderRadius: '6px', border: '1px solid #ccc', outline: 'none', fontSize: '1rem', resize: 'vertical' }}></textarea>
                    </div>

                    <button type="submit" style={{ backgroundColor: '#2a5fbd', color: 'white', padding: '14px', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '1.1rem', marginTop: '10px', transition: 'background 0.3s' }}>
                        Send Message <i className="fas fa-paper-plane" style={{ marginLeft: '8px' }}></i>
                    </button>
                </form>
            </div>
        </div>
    );
};

// --- About Us Component ---
const AboutUs = () => {
    return (
        <div style={{ padding: '60px 10%', flexGrow: 1, backgroundColor: '#ffffff', color: '#333', lineHeight: '1.8' }}>
            <div style={{ maxWidth: '900px', margin: '0 auto' }}>
                <h1 style={{ textAlign: 'center', color: '#1a2033', fontSize: '2.5rem', marginBottom: '30px', borderBottom: '3px solid #2a5fbd', paddingBottom: '10px', display: 'inline-block' }}>About Office Lite</h1>

                <p style={{ fontSize: '1.2rem', marginBottom: '20px', color: '#555' }}>
                    Welcome to <strong>Office Lite</strong>, your all-in-one web-based productivity suite designed for speed, simplicity, and privacy. We believe that creating professional documents shouldn't require heavy software or complex setups.
                </p>

                <h3 style={{ color: '#2a5fbd', marginTop: '30px' }}>Our Mission</h3>
                <p>Our mission is to provide a lightweight yet powerful platform where users can draft documents, design presentations, and manage data through spreadsheets—all directly within the browser. Whether you are a student, a professional, or a hobbyist, Office Lite is built to enhance your workflow.</p>

                <h3 style={{ color: '#2a5fbd', marginTop: '30px' }}>What We Offer</h3>
                <ul style={{ listStyleType: 'none', padding: 0 }}>
                    <li style={{ marginBottom: '15px' }}><strong><i className="fas fa-file-alt"></i> Word Editor:</strong> A clean interface for distraction-free writing and document formatting.</li>
                    <li style={{ marginBottom: '15px' }}><strong><i className="fas fa-project-diagram"></i> Presentation Slides:</strong> Simple tools to create impactful visual stories.</li>
                    <li style={{ marginBottom: '15px' }}><strong><i className="fas fa-table"></i> Spreadsheet Tools:</strong> Organize your data and perform calculations with ease.</li>
                </ul>

                <h3 style={{ color: '#2a5fbd', marginTop: '30px' }}>Privacy First</h3>
                <p>At Office Lite, we value your security. Unlike other platforms, we do not store your documents on our servers. Everything you create stays on your local device, giving you 100% control over your data.</p>

                <div style={{ marginTop: '40px', padding: '20px', backgroundColor: '#f8f9fa', borderRadius: '8px', textAlign: 'center', borderLeft: '5px solid #2a5fbd' }}>
                    <p style={{ margin: 0, fontWeight: 'bold' }}>"Simplifying productivity, one slide at a time."</p>
                </div>
            </div>
        </div>
    );
};

// --- Typewriter Effect Component ---
const TypewriterEffect = () => {
    const messages = [
        "Simple & Fast: Office Lite is a lightweight and fast platform to create documents, slides, and sheets effortlessly.",
        "Privacy First: Your data stays on your device; we never store or save your files on our servers.",
        "All-in-One Tools: From word editing to spreadsheets, all essential office tools are available right in your browser."
    ];

    const [displayText, setDisplayText] = React.useState("");
    const [messageIndex, setMessageIndex] = React.useState(0);
    const [charIndex, setCharIndex] = React.useState(0);
    const [isDeleting, setIsDeleting] = React.useState(false);

    React.useEffect(() => {
        const currentMessage = messages[messageIndex];
        const speed = isDeleting ? 30 : 70; // Type karne ki speed 70ms, delete karne ki 30ms

        const timer = setTimeout(() => {
            if (!isDeleting && charIndex < currentMessage.length) {
                // Type ho raha hai
                setDisplayText(currentMessage.substring(0, charIndex + 1));
                setCharIndex(prev => prev + 1);
            } else if (!isDeleting && charIndex === currentMessage.length) {
                // Poori line likhi gayi, ab 2 second ruko phir delete karo
                setTimeout(() => setIsDeleting(true), 2000);
            } else if (isDeleting && charIndex > 0) {
                // Delete ho raha hai
                setDisplayText(currentMessage.substring(0, charIndex - 1));
                setCharIndex(prev => prev - 1);
            } else if (isDeleting && charIndex === 0) {
                // Agli line par jao
                setIsDeleting(false);
                setMessageIndex((prev) => (prev + 1) % messages.length);
            }
        }, speed);

        return () => clearTimeout(timer);
    }, [charIndex, isDeleting, messageIndex]);

    return (
        <div style={{
            backgroundColor: '#f0f2f5',
            padding: '15px',
            textAlign: 'center',
            fontSize: '1.1rem',
            color: '#2a5fbd',
            fontWeight: '500',
            minHeight: '60px',
            borderBottom: '1px solid #ddd',
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center'
        }}>
            {/* Blink khatam karne ke liye simple span */}
            <span style={{ paddingRight: '5px' }}>
                {displayText}
            </span>
        </div>
    );
};

// --- Main App Component ---
const OfficeLiteApp = () => {
    const [currentView, setCurrentView] = React.useState('dashboard');
    const [isLegalOpen, setIsLegalOpen] = React.useState(false);
    const [isMenuOpen, setIsMenuOpen] = React.useState(false);

    const openEditor = (editorType) => {
        setCurrentView(editorType);
        setIsLegalOpen(false);
        setIsMenuOpen(false);
    };
    const closeEditor = () => { setCurrentView('dashboard'); setIsMenuOpen(false); };

    let mainBodyContent;

    if (currentView === 'dashboard') {
        mainBodyContent = (
            <section className="create-section" style={{ flexGrow: 1, display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
                <h2 className="create-title" style={{ textAlign: 'center', fontSize: '1.8rem', marginBottom: '40px' }}>CREATE NEW DOCUMENT</h2>
                <div className="category-cards" style={{ maxWidth: '1200px', margin: '0 auto', width: '100%', padding: '0 20px' }}>
                    {/* Word Card */}
                    <div className="category-card word-card" onClick={() => openEditor('word')}>
                        <div className="card-image-box"><i className="fas fa-file-word card-icon"></i></div>
                        <div className="card-content">
                            <h3 className="card-title">WORD</h3>
                            <p className="card-documents">Sample documents, reports, letters</p>
                            <button className="create-btn word-btn">Word Document</button>
                        </div>
                    </div>
                    {/* Presentation Card */}
                    <div className="category-card pres-card" onClick={() => openEditor('presentation')}>
                        <div className="card-image-box"><i className="fas fa-file-powerpoint card-icon"></i></div>
                        <div className="card-content">
                            <h3 className="card-title">PRESENTATION</h3>
                            <p className="card-documents">Slide decks, pitch presentations</p>
                            <button className="create-btn pres-btn">Presentation Slides</button>
                        </div>
                    </div>
                    {/* Sheet Card */}
                    <div className="category-card sheet-card" onClick={() => openEditor('sheet')}>
                        <div className="card-image-box"><i className="fas fa-file-excel card-icon"></i></div>
                        <div className="card-content">
                            <h3 className="card-title">SHEET</h3>
                            <p className="card-documents">Data analysis, spreadsheets, budgets</p>
                            <button className="create-btn sheet-btn">Spreadsheet</button>
                        </div>
                    </div>
                </div>
            </section>
        );
    } else if (currentView === 'word') {
        mainBodyContent = <WordEditor onClose={closeEditor} />;
    } else if (currentView === 'presentation') {
        mainBodyContent = <PresentationEditor onClose={closeEditor} />;
    } else if (currentView === 'sheet') {
        mainBodyContent = <SheetEditor onClose={closeEditor} />;
        // 👇 YE 4 LINES ADD KAREIN 👇
    } else if (currentView === 'privacy') {
        mainBodyContent = <PrivacyPolicy />;
    } else if (currentView === 'terms') {
        mainBodyContent = <TermsAndConditions />;
    } else if (currentView === 'contact') {
        mainBodyContent = <ContactUs />;
    } else if (currentView === 'about') {
        mainBodyContent = <AboutUs />;
    }

    return (
        <div className="dashboard-container" style={{ width: '100%', maxWidth: '100%', minHeight: '100vh', borderRadius: '0', display: 'flex', flexDirection: 'column' }}>

            {/* Header */}
            <header className="header">
                <div className="header-top" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', width: '100%' }}>
                    <div className="brand" onClick={closeEditor} style={{ cursor: 'pointer' }}>
                        <i className="fas fa-layer-group"></i> OFFICE LITE
                    </div>
                    <button className="hamburger-btn" onClick={() => setIsMenuOpen(!isMenuOpen)}>
                        <i className={`fas ${isMenuOpen ? 'fa-times' : 'fa-bars'}`}></i>
                    </button>
                </div>
                <nav className={`header-nav ${isMenuOpen ? 'nav-open' : ''}`}>
                    <span className={`nav-item ${currentView === 'dashboard' ? 'active' : ''}`} onClick={closeEditor} style={{ cursor: 'pointer' }}>HOME</span>
                    <span className={`nav-item ${currentView === 'word' ? 'active' : ''}`} onClick={() => openEditor('word')} style={{ cursor: 'pointer' }}>WORD</span>
                    <span className={`nav-item ${currentView === 'presentation' ? 'active' : ''}`} onClick={() => openEditor('presentation')} style={{ cursor: 'pointer' }}>PRESENTATION</span>
                    <span className={`nav-item ${currentView === 'sheet' ? 'active' : ''}`} onClick={() => openEditor('sheet')} style={{ cursor: 'pointer' }}>SHEET</span>

                    <div
                        style={{ position: 'relative' }}
                        onMouseEnter={() => setIsLegalOpen(true)}
                        onMouseLeave={() => setIsLegalOpen(false)}
                    >
                        <span
                            className={`nav-item ${currentView === 'privacy' || currentView === 'terms' ? 'active' : ''}`}
                            style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '5px' }}
                        >
                            LEGAL <i className="fas fa-caret-down"></i>
                        </span>

                        {/* Dropdown Menu */}
                        {isLegalOpen && (
                            <div style={{
                                position: 'absolute',
                                top: '100%',
                                left: 0,
                                backgroundColor: '#242d45',
                                borderRadius: '8px',
                                marginTop: '10px',
                                overflow: 'hidden',
                                zIndex: 1000,
                                minWidth: '200px',
                                boxShadow: '0 8px 16px rgba(0,0,0,0.3)',
                                border: '1px solid #3b456e'
                            }}>
                                <div
                                    onClick={() => openEditor('privacy')}
                                    style={{ padding: '12px 15px', color: 'white', cursor: 'pointer', borderBottom: '1px solid #3b456e', transition: 'background 0.2s' }}
                                    onMouseOver={(e) => e.target.style.backgroundColor = '#3b456e'}
                                    onMouseOut={(e) => e.target.style.backgroundColor = 'transparent'}
                                >
                                    <i className="fas fa-user-shield" style={{ width: '20px' }}></i> Privacy Policy
                                </div>
                                <div
                                    onClick={() => openEditor('terms')}
                                    style={{ padding: '12px 15px', color: 'white', cursor: 'pointer', transition: 'background 0.2s' }}
                                    onMouseOver={(e) => e.target.style.backgroundColor = '#3b456e'}
                                    onMouseOut={(e) => e.target.style.backgroundColor = 'transparent'}
                                >
                                    <i className="fas fa-file-contract" style={{ width: '20px' }}></i> Terms & Conditions
                                </div>

                            </div>

                        )}

                    </div>
                    <span className={`nav-item ${currentView === 'contact' ? 'active' : ''}`} onClick={() => openEditor('contact')} style={{ cursor: 'pointer' }}>CONTACT US</span>
                    <span className={`nav-item ${currentView === 'about' ? 'active' : ''}`} onClick={() => openEditor('about')} style={{ cursor: 'pointer' }}>ABOUT</span>

                </nav>

            </header>

            {/* 👇 YAHAN TYPEWRITER ADD KAREIN (Sirf Dashboard view par dikhega) 👇 */}
            {currentView === 'dashboard' && <TypewriterEffect />}

            {/* Main Content */}
            <div style={{ flexGrow: 1, display: 'flex', flexDirection: 'column' }}>
                {mainBodyContent}
            </div>


            {/* Footer */}
            <footer className="footer" style={{
                marginTop: 'auto',
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                gap: '8px',
                padding: '30px 0',
                background: '#1a2033',
                color: 'white'
            }}>
                <span style={{ fontSize: '1.4rem', fontWeight: 'bold', letterSpacing: '2px' }}>OFFICE LITE</span>
                <span style={{ fontSize: '1rem', color: '#adb5bd' }}>
                    Created with <i className="fas fa-heart" style={{ color: '#e74c3c' }}></i> by Paras
                </span>
                <span style={{ fontSize: '0.8rem', color: '#6c757d', marginTop: '5px' }}>
                    &copy; 2026 All Rights Reserved.
                </span>
            </footer>

        </div>
    );
};

// Render the application
ReactDOM.render(<OfficeLiteApp />, document.getElementById('root'));