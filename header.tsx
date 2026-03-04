import * as React from 'react';
import { Link, Image, makeStyles, shorthands } from '@fluentui/react-components';

const logo: string = require('../img/AMFCNewLogo.png');

const useStyles = makeStyles({
  wrapper: {
    backgroundColor: '#ffffff',
    ...shorthands.border('1px', 'solid', '#d0d0d0'),
    ...shorthands.padding('20px', '24px'),
    maxWidth: '760px',
    marginLeft: 'auto',
    marginRight: 'auto',
    ...shorthands.margin('16px', '0'),
    boxShadow: '0 2px 10px rgba(0,0,0,0.12)',
  },

  content: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    textAlign: 'center',
    rowGap: '10px',
  },

  title: {
    fontSize: '26px',
    fontWeight: 700,
    color: '#1a3f7a',
  },

  subtitle: {
    fontSize: '14px',
    color: '#333',
    maxWidth: '680px',
  },

  note: {
    fontSize: '13px',
    fontWeight: 600,
    color: '#b10000',
  },

  logo: {
    width: '220px',
  },
});

const HeaderComponent = (): JSX.Element => {
  const styles = useStyles();

  return (
    <div className={styles.wrapper}>
      <div className={styles.content}>

        <Image src={logo} alt="AmeriHealth Caritas" className={styles.logo} />

        <div className={styles.title}>
          Let’s Fix It
        </div>

        <div className={styles.subtitle}>
          Please use this form to submit any issues identified while navigating in the new Online Help environment.
        </div>

        <div className={styles.subtitle}>
          <Link
            href="https://amerihealthcaritas.sharepoint.com/sites/eokm/Online%20Help%20Assets/TopicView.aspx?ID=Let%26%23039%3Bs%20Fix%20It-LstNme=Systems"
            target="_blank"
          >
            Please refer to the Online Help Topic “Let’s Fix It” for further instructions.
          </Link>
        </div>

        <div className={styles.note}>
          Note: * Red asterisk indicates a required field
        </div>

      </div>
    </div>
  );
};

export default HeaderComponent;