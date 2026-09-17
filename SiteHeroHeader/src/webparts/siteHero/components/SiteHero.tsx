import * as React from "react";
import styles from "./SiteHero.module.scss";
import { ISiteHeroProps } from "./ISiteHeroProps";

const SiteHero: React.FC<ISiteHeroProps> = ({
  heroText,
  subTitle,
  backgroundImageUrl
}) => {
  const backgroundStyle: React.CSSProperties = backgroundImageUrl
    ? {
        backgroundImage: `url("${backgroundImageUrl}")`
      }
    : {};

  return (
    <section
      className={styles.header}
      style={backgroundStyle}
      dir="rtl"
    >
      <div className={styles.container}>
        <div className={styles.textGroup}>
          {heroText ? (
            <h1 className={styles.title}>{heroText}</h1>
          ) : null}

          {subTitle ? (
            <p className={styles.subtitle}>{subTitle}</p>
          ) : null}
        </div>
      </div>
    </section>
  );
};

export default SiteHero;