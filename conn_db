#conn_db
import os
from sqlalchemy import create_engine, Date, String, Integer, Column
from sqlalchemy.orm import declarative_base, sessionmaker
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'mits.db')

engine = create_engine(f'sqlite:///{DB_PATH}', echo=True)

Base = declarative_base()

class VerificationRecord(Base):
    __tablename__ = 'verifications'

    id = Column(Integer, primary_key=True, autoincrement=True)
    act = Column(String, nullable=False, comment='Акт')
    instrument_name = Column(String, nullable=False, comment='Наименование СИ')
    serial_number = Column(String, nullable=False, comment='Заводской №')
    verification_date = Column(Date, nullable=False, comment='Дата поверки')
    next_verification_date = Column(Date, nullable=True, comment='Дата след. поверки')
    note = Column(String, nullable=True, comment='Примечание')
    mit_link = Column(String, nullable=True, comment='Ссылка на протокол')

    def __repr__(self):
        return (f"<VerificationRecord(act='{self.act}', "
                f"instrument='{self.instrument_name}', "
                f"serial='{self.serial_number}', "
                f"verif_date={self.verification_date}, "
                f"next_date={self.next_verification_date})>")


        

Base.metadata.create_all(engine)
Session=sessionmaker(bind=engine)


def add_mits(data: list):
    with Session() as session:
        mits = []
        for el in data:
            next_date = None
            if el.get('next_verification_date'):
                next_date = datetime.strptime(el.get('next_verification_date'), '%Y-%m-%d').date()
            mits.append({
                    'act': el.get('act'),
                    'instrument_name': el.get('instrument_name'),
                    'serial_number': el.get('serial_number'),
                    'verification_date': datetime.strptime(el.get('verification_date'), '%Y-%m-%d').date(),
                    'next_verification_date':  next_date,
                    'note': el.get('note'),
                    'mit_link': el.get('mit_link')
                })
        session.bulk_insert_mappings(VerificationRecord, mits)
        session.commit()


