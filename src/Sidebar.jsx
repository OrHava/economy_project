import React, { useState } from "react";
import { FiMenu, FiX } from "react-icons/fi";
import "./Sidebar.css";

function Sidebar() {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <div className={`sidebar ${isOpen ? "open" : ""}`}>
      <div className="menu-toggle" onClick={() => setIsOpen(!isOpen)}>
        {isOpen ? <FiX size={24} /> : <FiMenu size={24} />}
      </div>

      <nav className="nav-menu">
        <h1 className="site-title">Elevate</h1>
        <ul className="nav-list">
          {["Home", "About", "Services", "Contact"].map((item) => (
            <li key={item} className="nav-item" style={{ display: "block", marginBottom: "1rem" }}>
              <a href={`#${item.toLowerCase()}`} className="nav-link">
                {item}
              </a>
            </li>
          ))}
        </ul>
      </nav>
    </div>
  );
}

export default Sidebar;
