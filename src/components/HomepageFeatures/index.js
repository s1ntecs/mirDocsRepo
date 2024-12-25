import clsx from 'clsx';
import Heading from '@theme/Heading';
import styles from './styles.module.css';

const FeatureList = [
  {
    title: 'Быстрое использование',
    Img: require('@site/static/img/fastUse.webp').default, // Изменено на Img
    description: (
      <>
        Для использования ПО нужен только Excel и вы можете строить трубопроводные сети.
      </>
    ),
  },
  {
    title: 'Удобный и понятный интерфейс',
    Img: require('@site/static/img/interface.png').default, // Изменено на Img
    description: (
      <>
        У ПО удобный и понятный интерфейс, соединяйте объекты (геометрические фигуры) с линиями (трубопроводами).
      </>
    ),
  },
  {
    title: 'Большой функционал',
    Img: require('@site/static/img/functions.webp').default, // Изменено на Img
    description: (
      <>
        У ПО Мир есть разный функционал для слежения за потоками флюида.
      </>
    ),
  },
];

function Feature({Img, title, description}) {
  return (
    <div className={clsx('col col--4')}>
      <div className="text--center">
        <img src={Img} alt={title} className={styles.featureImg} role="img" />
      </div>
      <div className="text--center padding-horiz--md">
        <Heading as="h3">{title}</Heading>
        <p>{description}</p>
      </div>
    </div>
  );
}

export default function HomepageFeatures() {
  return (
    <section className={styles.features}>
      <div className="container">
        <div className="row">
          {FeatureList.map((props, idx) => (
            <Feature key={idx} {...props} />
          ))}
        </div>
      </div>
    </section>
  );
}
