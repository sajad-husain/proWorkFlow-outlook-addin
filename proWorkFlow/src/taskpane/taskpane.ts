import * as React from 'react';
import * as ReactDOM from 'react-dom';
import App from './components/App';

/* Render application after Office initializes */
Office.onReady(() => {
  ReactDOM.render(
    React.createElement(App),
    document.getElementById('container')
  );
});