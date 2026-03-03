/**
 * HeaderComponent.tsx
 *
 * Uses Fluent UI v9 components to render a standard “form header” block:
 * - AmeriHealth Caritas logo
 * - Title line
 * - Subtitle line
 * - Note line (red) explaining required fields
 *
 * This is intended to be a simple, reusable header you can drop at the top of any form.
 *
 * Example usage:
 *
 * <HeaderComponent
 *   title="Form Title Goes Here"
 *   subtitle="Form Subtitle goes here"
 * />
 *
 * // With a help link line:
 * <HeaderComponent
 *   title="Let's Fix It"
 *   subtitle="Please use this form to submit any issues identified while navigating in new Online Help Environment."
 *   linkUrl="https://amerihealthcaritas.sharepoint.com/sites/eokm/Online%20Help%20Assets/TopicView.aspx?ID=Let%26%23039%3Bs%20Fix%20It-LstNme=Systems"
 *   linkText="Please refer to the Online Help Topic “Let’s Fix It” for further instructions."
 * />
 */

/**
 * HeaderComponent.tsx
 *
 * Renders a standard form header block:
 * - AmeriHealth Caritas logo
 * - Title
 * - Subtitle
 * - Optional help link line
 * - Required-fields note (red)
 *
 * Example usage:
 *
 * <HeaderComponent
 *   title="Form Title Goes Here"
 *   subtitle="Form Subtitle goes here"
 * />
 *
 * <HeaderComponent
 *   title="Let's Fix It"
 *   subtitle="Please use this form to submit any issues identified while navigating in new Online Help Environment."
 *   linkUrl="https://amerihealthcaritas.sharepoint.com/sites/eokm/Online%20Help%20Assets/TopicView.aspx?ID=Let%26%23039%3Bs%20Fix%20It-LstNme=Systems"
 *   linkText="Please refer to the Online Help Topic “Let’s Fix It” for further instructions."
 * />
 */

import * as React from 'react';
import { Link, Image, makeStyles, shorthands } from '@fluentui/react-components';

// NOTE: Path must be correct relative to THIS file.
// Based on your explorer, the image is under: src/extensions/letsFixIt/img/AMFCNewLogo.png
// This component is under: src/extensions/letsFixIt/components/HeaderComponent.tsx
// So the correct relative path is: ../img/AMFCNewLogo.png
import AMFCNewLogo from '../img/AMFCNewLogo.png';

export interface HeaderComponentProps {
  /** Main header title (large, bold, centered) */
  title: string;

  /** Supporting subtitle (smaller, centered) */
  subtitle?: string;

  /**
   * Optional “help / instructions” hyperlink line.
   * If linkUrl is provided, linkText will be shown as the clickable label.
   */
  linkUrl?: string;
  linkText?: string;

  /**
   * Optional override for the required-fields note line.
   * Default: "Note: * Red asterisk indicates a required field"
   */
  noteText?: string;

  /** Optional wrapper className if you want to control spacing externally */
  className?: string;
}

const DEFAULT_NOTE = 'Note: * Red asterisk indicates a required field';

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
    lineHeight: 1.2,
  },

  subtitle: {
    fontSize: '14px',
    color: '#333333',
    maxWidth: '680px',
    lineHeight: 1.35,
  },

  linkLine: {
    fontSize: '13px',
    color: '#333333',
    maxWidth: '680px',
    lineHeight: 1.35,
  },

  note: {
    fontSize: '13px',
    fontWeight: 600,
    color: '#b10000',
  },

  logo: {
    width: '220px',
    height: 'auto',
  },
});

const HeaderComponent = (props: HeaderComponentProps): JSX.Element => {
  const { title, subtitle, linkUrl, linkText, noteText = DEFAULT_NOTE, className } = props;

  const styles = useStyles();

  // If linkUrl is provided but no linkText is provided, fall back to showing the URL.
  const resolvedLinkText: string | undefined = linkUrl ? (linkText || linkUrl) : undefined;

  return (
    <div className={`${styles.wrapper}${className ? ` ${className}` : ''}`}>
      <div className={styles.content}>
        {/* Logo */}
        <Image src={AMFCNewLogo} alt="AmeriHealth Caritas" className={styles.logo} />

        {/* Title */}
        <div className={styles.title}>{title}</div>

        {/* Subtitle (optional) */}
        {subtitle ? <div className={styles.subtitle}>{subtitle}</div> : null}

        {/* Help link (optional) */}
        {linkUrl ? (
          <div className={styles.linkLine}>
            <Link href={linkUrl} target="_blank" rel="noreferrer">
              {resolvedLinkText}
            </Link>
          </div>
        ) : null}

        {/* Required note */}
        <div className={styles.note}>{noteText}</div>
      </div>
    </div>
  );
};

export default HeaderComponent;