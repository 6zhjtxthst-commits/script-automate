import React, { useState } from "react";
import "./styles.css";

export default function App() {
  const [count, setCount] = useState(0);

  return (
    <div className="page-shell">
      <header className="topbar">
        <a className="brand" href="#home" aria-label="Little Things home">
          <span className="brand-mark" aria-hidden="true">✳</span>
          <span>little things<span className="brand-period">.</span></span>
        </a>
        <span className="topbar-note">A tiny React experiment <span aria-hidden="true">↗</span></span>
      </header>

      <main id="home" className="hero">
        <section className="intro" aria-labelledby="page-title">
          <div className="eyebrow"><span className="eyebrow-line" /> THE CLICK COUNTER · NO. 01</div>
          <h1 id="page-title">Make the<br /><em>little things</em><br />count.</h1>
          <p>Sometimes all you need is a little momentum. Start with one click, then see where it takes you.</p>
          <div className="intro-footer">
            <span className="asterisk" aria-hidden="true">✳</span>
            <span>Small actions.<br />Good energy.</span>
          </div>
        </section>

        <section className="counter-card" aria-label="Click counter">
          <div className="card-topline">
            <span>YOUR LITTLE MOMENT</span>
            <span className="card-sparkle" aria-hidden="true">✦</span>
          </div>

          <div className="count-display" aria-live="polite" aria-atomic="true">
            <span className="count-number" key={count}>{count}</span>
            <span className="count-label">{count === 1 ? "click so far" : "clicks so far"}</span>
          </div>

          <div className="card-bottom">
            <button className="add-button" type="button" onClick={() => setCount((value) => value + 1)}>
              <span>Add a click</span>
              <span className="button-arrow" aria-hidden="true">↗</span>
            </button>
            <button className="reset-button" type="button" onClick={() => setCount(0)} disabled={count === 0}>
              Start over <span aria-hidden="true">↺</span>
            </button>
          </div>
        </section>
      </main>

      <footer className="page-footer">
        <span>ONE CLICK AT A TIME</span>
        <span>MADE FOR THE JOY OF IT <span aria-hidden="true">✳</span></span>
      </footer>
    </div>
  );
}
