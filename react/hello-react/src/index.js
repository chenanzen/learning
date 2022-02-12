import React, {useState} from 'react';
import ReactDOM from 'react-dom';
import './index.css';

let city = { name: "singapore", country: "singapore" };

function Lake({name}){
  return <h1>Lake {name}!</h1>
}

function SkiResport(){
  return <h1>Ski Report!!</h1>
}

function App(props){
  const [status, setStatus] = useState("Open");
  return (
    <>
      <h1>Status:{status}</h1>
      <button onClick={() => }>Open</button>
    </>
  )
}

ReactDOM.render(
    <App name="react" />,
    document.getElementById('root')
  );


const snacks = ["Popcorn", "Chips", "Cassava"];