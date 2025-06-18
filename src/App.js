import React from 'react';
import './App.css';
import { BrowserRouter as Router, Route, Switch, Redirect } from 'react-router-dom';
import Luckysheet from './component/Luckysheet';

function App() {
  return (
    <div className="App">
      <Router>
        <Switch>
          <Route exact path="/" component={Luckysheet} />
        </Switch>
      </Router>
    </div>
  );
}

export default App;
