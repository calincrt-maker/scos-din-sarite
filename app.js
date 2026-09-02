document.addEventListener('DOMContentLoaded', () => {
  const data = window.presentationData;
  if (!data) return;

  const totalSlides = data.slides.length;
  const viewportContainer = document.getElementById('viewportContainer');
  const slideCounter = document.getElementById('slideCounter');
  const progressFill = document.getElementById('progressFill');
  const sideDotsNav = document.getElementById('sideDotsNav');
  const overviewModal = document.getElementById('overviewModal');
  const overviewGrid = document.getElementById('overviewGrid');

  let currentActiveIndex = 0;

  // BUILD ALL SLIDES VERTICALLY IN THE SCROLL CONTAINER
  function buildAllSlides() {
    viewportContainer.innerHTML = '';
    sideDotsNav.innerHTML = '';

    data.slides.forEach((slide, idx) => {
      // 1. Create section container
      const section = document.createElement('section');
      section.className = 'slide-section';
      section.id = `slide-${idx}`;
      section.setAttribute('data-index', idx);

      let bodyHTML = '';

      switch (slide.layout) {
        case 'cover':
          bodyHTML = `
            <div class="layout-cover">
              <div class="cover-left">
                <div class="slide-tag">${slide.tag}</div>
                <h1 class="cover-title">${slide.title}</h1>
                <p class="cover-desc">${slide.desc}</p>
                <div class="cover-pills">
                  ${slide.metaPills.map(p => `<span class="cover-pill">${p}</span>`).join('')}
                </div>
              </div>
              <div class="cover-right">
                <div class="cover-img-wrapper">
                  <img src="${slide.bgImage}" alt="${slide.title}">
                </div>
              </div>
            </div>
          `;
          break;

        case 'combined-context-concept':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="quote-banner-small" style="background:rgba(230,20,50,0.1); border-left:3px solid var(--red-neon); padding:10px 16px; border-radius:4px; font-family:var(--font-heading); font-size:14px; font-weight:700; color:var(--text-white); margin-bottom:14px;">
                ${slide.quote}
              </div>
              <div style="display:grid; grid-template-columns: 1fr 1.2fr; gap:20px; align-items:stretch;">
                <div class="card-box" style="padding:18px;">
                  <span class="card-badge">${slide.columns[0].badge}</span>
                  <h3 class="card-title" style="font-size:16px; margin-bottom:8px;">${slide.columns[0].title}</h3>
                  <p class="card-text" style="font-size:12px; line-height:1.45;">${slide.columns[0].text}</p>
                </div>
                <div style="display:flex; flex-direction:column; gap:10px;">
                  ${slide.pillars.map(p => `
                    <div class="pillar-card-compact" style="background:var(--bg-card); border:1px solid var(--border-card); border-radius:var(--radius-sm); padding:10px 14px; display:flex; gap:12px; align-items:center;">
                      <div style="font-family:var(--font-heading); font-size:20px; font-weight:700; color:var(--red-neon); flex-shrink:0;">${p.num}</div>
                      <div>
                        <div style="font-family:var(--font-heading); font-size:13px; font-weight:700; color:var(--text-white); margin-bottom:2px;">${p.title}</div>
                        <div style="font-size:11px; color:var(--text-secondary); line-height:1.3;">${p.desc}</div>
                      </div>
                    </div>
                  `).join('')}
                </div>
              </div>
            </div>
          `;
          break;

        case 'feature-split':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="layout-feature-split">
                <div class="feature-list">
                  ${slide.features.map(f => `
                    <div class="feature-item">
                      <div class="feature-item-title">${f.title}</div>
                      <div class="feature-item-desc">${f.desc}</div>
                    </div>
                  `).join('')}
                </div>
                <div class="feature-img-frame">
                  <img src="${slide.image}" alt="${slide.title}">
                </div>
              </div>
            </div>
          `;
          break;

        case 'portrait-spotlight':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="layout-portrait">
                <div class="portrait-img-box">
                  <img src="${slide.image}" alt="${slide.title}">
                </div>
                <div>
                  <p class="card-text" style="margin-bottom:14px;">${slide.desc}</p>
                  <div class="highlights-grid">
                    ${slide.highlights.map(h => `<div class="highlight-chip">${h}</div>`).join('')}
                  </div>
                </div>
              </div>
            </div>
          `;
          break;

        case 'tri-cards':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="layout-tri-cards">
                ${slide.cards.map(c => `
                  <div class="tri-card">
                    <div class="tri-stat">${c.stat}</div>
                    <div class="tri-sub">${c.sub}</div>
                    <div class="tri-body">${c.body}</div>
                  </div>
                `).join('')}
              </div>
            </div>
          `;
          break;

        case 'stats-grid':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="stats-row">
                ${slide.bigStats.map(s => `
                  <div class="stat-item-box">
                    <div class="stat-item-num">${s.num}</div>
                    <div class="stat-item-label">${s.label}</div>
                    <div class="stat-item-detail">${s.detail}</div>
                  </div>
                `).join('')}
              </div>
              <div class="schedule-card">
                ${slide.schedule.map(sch => `
                  <div class="schedule-row">
                    <div class="sch-slot">${sch.slot}</div>
                    <div class="sch-time">${sch.time}</div>
                    <div class="sch-note">${sch.note}</div>
                  </div>
                `).join('')}
              </div>
            </div>
          `;
          break;

        case 'gallery-3col':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="layout-gallery">
                ${slide.gallery.map(g => `
                  <div class="gallery-card">
                    <div class="gallery-img-box">
                      <img src="${g.image}" alt="${g.title}">
                    </div>
                    <div class="gallery-info">
                      <div class="gallery-title">${g.title}</div>
                      <div class="gallery-cap">${g.caption}</div>
                    </div>
                  </div>
                `).join('')}
              </div>
            </div>
          `;
          break;

        case 'brand-pillars':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="layout-brand-grid">
                ${slide.grid.map(b => `
                  <div class="brand-card">
                    <div class="brand-icon">${b.icon}</div>
                    <div>
                      <div class="brand-title-small">${b.title}</div>
                      <div class="brand-desc-small">${b.desc}</div>
                    </div>
                  </div>
                `).join('')}
              </div>
            </div>
          `;
          break;

        case 'ecosystem-4col':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="layout-eco-grid">
                ${slide.elements.map(e => `
                  <div class="eco-card">
                    <div class="eco-code">${e.code}</div>
                    <div class="eco-name">${e.name}</div>
                    <div class="eco-sec">${e.sec}</div>
                    <div class="eco-desc">${e.desc}</div>
                  </div>
                `).join('')}
              </div>
            </div>
          `;
          break;

        case 'media-table':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="table-wrapper">
                <table class="media-table">
                  <thead>
                    <tr>
                      ${slide.table.headers.map(h => `<th>${h}</th>`).join('')}
                    </tr>
                  </thead>
                  <tbody>
                    ${slide.table.rows.map(row => `
                      <tr>
                        ${row.map(cell => {
                          if (cell.includes('BONUS')) {
                            return `<td><span class="bonus-tag">${cell}</span></td>`;
                          }
                          return `<td>${cell}</td>`;
                        }).join('')}
                      </tr>
                    `).join('')}
                  </tbody>
                </table>
                <div class="table-summary-bar">${slide.table.total}</div>
              </div>
            </div>
          `;
          break;

        case 'pricing-card':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="layout-pricing">
                <div class="pricing-hero-card">
                  <div class="card-badge">PACHET SEZON II</div>
                  <div class="pricing-num">${slide.commercial.totalNet}</div>
                  <div class="pricing-sub">${slide.commercial.currencyNote}</div>
                  <div class="pricing-weekly">
                    <div class="card-badge" style="margin-bottom:2px;">ECHIVALENT SĂPTĂMÂNAL</div>
                    <div class="pricing-weekly-val">${slide.commercial.weeklyNet}</div>
                  </div>
                </div>
                <div class="pricing-inc-card">
                  <div class="card-badge" style="margin-bottom:10px;">INCLUSIVITATE & GARANȚII</div>
                  <div class="inc-list">
                    ${slide.commercial.inclusions.map(inc => `<div class="inc-item">${inc}</div>`).join('')}
                  </div>
                </div>
              </div>
            </div>
          `;
          break;

        case 'emotion-quote':
          bodyHTML = `
            <div class="slide-tag">${slide.tag}</div>
            <h2 class="slide-header-title">${slide.title}</h2>
            <p class="slide-header-sub">${slide.subtitle}</p>
            <div class="slide-body">
              <div class="layout-quote-emotion">
                <div>
                  <h3 class="big-quote-text">${slide.quote}</h3>
                  <p class="quote-desc-text">${slide.desc}</p>
                </div>
                <div class="cover-img-wrapper" style="height:300px;">
                  <img src="${slide.image}" alt="${slide.title}">
                </div>
              </div>
            </div>
          `;
          break;

        case 'closing':
          bodyHTML = `
            <div class="layout-closing">
              <div class="slide-tag">${slide.tag}</div>
              <h1 class="closing-title">${slide.title}</h1>
              <p class="closing-sub">${slide.callToAction}</p>
              
              <div class="contact-card-box">
                <div class="card-badge" style="margin-bottom:4px;">PERSOANĂ DE CONTACT & PARTENERIATE</div>
                <div style="font-family:var(--font-heading); font-size:22px; font-weight:700; color:var(--text-white); margin-bottom:8px;">
                  ${slide.contact.contactPerson}
                </div>
                <div style="display:flex; gap:16px; font-size:14px; font-weight:600; color:var(--red-soft); margin-bottom:12px;">
                  <a href="mailto:${slide.contact.email}" style="color:var(--red-soft); text-decoration:none;">📧 ${slide.contact.email}</a>
                  <span>•</span>
                  <a href="tel:${slide.contact.phone.replace(/\s/g, '')}" style="color:var(--text-white); text-decoration:none;">📞 ${slide.contact.phone}</a>
                </div>
                <div class="contact-pub" style="font-size:12px; color:var(--text-secondary);">${slide.contact.publisher} — ${slide.contact.address}</div>
                <div class="contact-detail" style="font-size:11px; color:var(--text-muted);">${slide.contact.web} • ${slide.contact.fb}</div>
              </div>
            </div>
          `;
          break;
      }

      section.innerHTML = `
        <div class="slide-stage">
          <div class="slide-content active">
            ${bodyHTML}
          </div>
        </div>
      `;
      viewportContainer.appendChild(section);

      // 2. Create side dot indicator
      const dot = document.createElement('div');
      dot.className = `dot-item ${idx === 0 ? 'active' : ''}`;
      dot.title = `Slide ${idx + 1}: ${slide.title}`;
      dot.addEventListener('click', () => scrollToSlide(idx));
      sideDotsNav.appendChild(dot);
    });
  }

  // SCROLL TO SPECIFIC SLIDE BY INDEX
  function scrollToSlide(idx) {
    if (idx < 0) idx = 0;
    if (idx >= totalSlides) idx = totalSlides - 1;
    const targetSection = document.getElementById(`slide-${idx}`);
    if (targetSection) {
      targetSection.scrollIntoView({ behavior: 'smooth' });
    }
  }

  // INTERSECTION OBSERVER TO TRACK CURRENT VISIBLE SLIDE ON SCROLL
  const observerOptions = {
    root: viewportContainer,
    threshold: 0.5
  };

  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const idx = parseInt(entry.target.getAttribute('data-index'), 10);
        currentActiveIndex = idx;
        updateActiveState(idx);
      }
    });
  }, observerOptions);

  function updateActiveState(index) {
    slideCounter.textContent = `SLIDE ${String(index + 1).padStart(2, '0')} / ${String(totalSlides).padStart(2, '0')}`;
    const pct = ((index + 1) / totalSlides) * 100;
    progressFill.style.width = `${pct}%`;

    // Update side dots
    const dots = sideDotsNav.querySelectorAll('.dot-item');
    dots.forEach((dot, dIdx) => {
      dot.classList.toggle('active', dIdx === index);
    });

    // Update overview modal
    updateOverviewActiveState();
  }

  // NAVIGATION BUTTONS
  document.getElementById('btnNext').addEventListener('click', () => {
    scrollToSlide(currentActiveIndex + 1);
  });
  document.getElementById('btnPrev').addEventListener('click', () => {
    scrollToSlide(currentActiveIndex - 1);
  });

  // OVERVIEW GRID MODAL
  function renderOverviewGrid() {
    overviewGrid.innerHTML = data.slides.map((s, idx) => `
      <div class="thumb-card ${idx === currentActiveIndex ? 'active' : ''}" data-index="${idx}">
        <span class="thumb-num">SLIDE ${String(idx + 1).padStart(2, '0')}</span>
        <div class="thumb-title">${s.title}</div>
      </div>
    `).join('');

    overviewGrid.querySelectorAll('.thumb-card').forEach(card => {
      card.addEventListener('click', (e) => {
        const idx = parseInt(e.currentTarget.getAttribute('data-index'), 10);
        scrollToSlide(idx);
        closeOverviewModal();
      });
    });
  }

  function updateOverviewActiveState() {
    overviewGrid.querySelectorAll('.thumb-card').forEach((card, idx) => {
      card.classList.toggle('active', idx === currentActiveIndex);
    });
  }

  function openOverviewModal() {
    renderOverviewGrid();
    overviewModal.classList.add('active');
  }

  function closeOverviewModal() {
    overviewModal.classList.remove('active');
  }

  document.getElementById('btnOverview').addEventListener('click', () => {
    if (overviewModal.classList.contains('active')) {
      closeOverviewModal();
    } else {
      openOverviewModal();
    }
  });

  document.getElementById('btnCloseOverview').addEventListener('click', closeOverviewModal);

  document.getElementById('btnFullscreen').addEventListener('click', () => {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(err => console.log(err));
    } else {
      document.exitFullscreen();
    }
  });

  document.getElementById('btnPrintPdf').addEventListener('click', () => {
    window.print();
  });

  document.getElementById('btnDownloadPptx').addEventListener('click', () => {
    const link = document.createElement('a');
    link.href = 'Scos_din_sarite_cu_Serghei_Sezonul_2_Pitch_Deck.pptx';
    link.download = 'Scos_din_sarite_cu_Serghei_Sezonul_2_Pitch_Deck.pptx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  });

  // KEYBOARD NAVIGATION (UP/DOWN & LEFT/RIGHT & SPACE)
  window.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowRight' || e.key === 'Space' || e.key === 'PageDown') {
      e.preventDefault();
      scrollToSlide(currentActiveIndex + 1);
    } else if (e.key === 'ArrowUp' || e.key === 'ArrowLeft' || e.key === 'PageUp') {
      e.preventDefault();
      scrollToSlide(currentActiveIndex - 1);
    } else if (e.key === 'Escape') {
      closeOverviewModal();
    } else if (e.key === 'g' || e.key === 'G') {
      if (overviewModal.classList.contains('active')) {
        closeOverviewModal();
      } else {
        openOverviewModal();
      }
    } else if (e.key === 'f' || e.key === 'F') {
      document.getElementById('btnFullscreen').click();
    }
  });

  // INITIALIZE
  buildAllSlides();
  document.querySelectorAll('.slide-section').forEach(sec => observer.observe(sec));
  updateActiveState(0);
});
