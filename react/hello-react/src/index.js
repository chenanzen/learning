import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';

let city = { name: "singapore", country: "singapore" };

ReactDOM.render(
    <h1 id="heading" className='cooltext'>{city.name} is in {city.country}</h1>,
    document.getElementById('root')
  );
